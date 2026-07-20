import { SC, LAND_NUMBER, GameMode, T, A } from './constants.js';
import { state } from './state.js';
import { abilityEffect } from './engine.js';

const rint = n => Math.floor(n);
// [di, dj, directionCode]. The engine encodes direction as a product of primes
// (2=up, 3=down, 5=right, 7=left), so a diagonal is the product of its two
// cardinal codes: up+right=2*5=10, down+right=3*5=15, down+left=3*7=21,
// up+left=2*7=14. The engine moves the player diagonally in one tick and its
// collision check (positionCheck) inspects ONLY the destination tile — there is
// no corner-cut restriction, so squeezing diagonally between two wall corners is
// legal. Navigating 8-connected lets the AI traverse corner-only connections
// that a 4-connected field reports as unreachable (the "詰む" cases).
const DIRS = [
  [0,-1,2], [0,1,3], [1,0,5], [-1,0,7],     // cardinal
  [1,-1,10], [1,1,15], [-1,1,21], [-1,-1,14], // diagonal
];

// ── Bookkeeping ──────────────────────────────────────────────────────────────
// Cycle detection by "progress signature" rather than raw position.
// Walking into a monster/box/wall restores the player's position (engine.js
// positionCheck), so a position-only check counts every combat tick as stuck.
// The signature includes monster HP and player HP/EXP, so any productive
// combat shows up as forward progress and resets the stuck counter.

const HISTORY_SIZE = 8;
let _sigHistory = [];
let _stuckTicks = 0;
let _lastFloor = -1;
let _floorOpenedHandled = false;
let _prevI = -1, _prevJ = -1; // player tile on the previous decision tick
const STUCK_THRESHOLD = 4;

// Auto-retry: ticks to linger on the GameOver screen before restarting (~2s).
const RETRY_DELAY_TICKS = 40;
let _retryDelay = 0;

// Drive auto-retry through the SAME keydown path a human uses by dispatching a
// real keyboard event, rather than calling the engine directly (which proved
// unreliable in the browser — the GameOver screen would never advance). The
// transitions GameOver --KeyX--> Entrance --Enter--> Dungeon all live in
// input.js's keydown handler, so this is identical to the player pressing them.
function pressKey(code) {
  if (typeof window !== 'undefined' && typeof KeyboardEvent !== 'undefined') {
    window.dispatchEvent(new KeyboardEvent('keydown', { code }));
  }
}

function resetCycleState() {
  _sigHistory = [];
  _stuckTicks = 0;
  _prevI = -1;
  _prevJ = -1;
  _lockIdx = -1;
}

function progressSig(pi, pj) {
  const { player, abilityHp } = state;
  // Sum of all alive monsters' HP — drops when player attacks, KeyZ damages,
  // or a monster dies. The single best progress signal during combat.
  let monsterHpSum = 0;
  for (let i = 0; i < state.monsterNumber; i++) {
    const m = state.monsters[i];
    if (m.alive) monsterHpSum += m.hp;
  }
  return `${pi},${pj}|${rint(player.hp)}|${player.exp}|${state.floor}|${abilityHp.join(',')}|${monsterHpSum}`;
}

function updateCycleDetection(sig) {
  const inCycle = _sigHistory.includes(sig);
  if (inCycle) {
    _stuckTicks++;
  } else {
    _stuckTicks = 0;
  }
  _sigHistory.push(sig);
  if (_sigHistory.length > HISTORY_SIZE) _sigHistory.shift();
}

function countAliveMonsters() {
  let n = 0;
  for (let i = 0; i < state.monsterNumber; i++) if (state.monsters[i].alive) n++;
  return n;
}

// ── Navigation: weighted cost field to the stair ─────────────────────────────
// Earlier versions stepped to the lowest-distance neighbour using a plain BFS
// and applied per-step penalties (avoid enemies, avoid red boxes) only when
// COMPARING the immediate neighbours. That created local minima: e.g. a stair
// reachable only by smashing a red box that sits at the mouth of a dead-end
// corridor — the +penalty made breaking the box look worse than walking up the
// (dead-end) corridor, so the player ping-ponged forever ("右往左往").
//
// The fix is to bake every cost (breaking a box, the red-box aversion, fighting
// a monster) into a Dijkstra distance field seeded from the stair. The result
// is a true potential: every tile has a neighbour with strictly lower field
// value that leads to the stair, so simply descending it can never oscillate
// and never gets trapped behind a "expensive but necessary" tile.
//
// CRITICAL: on the real maps corridors are routinely plugged with BOXES the
// player must smash through (breaking a box clears its tile to floor), so boxes
// must be traversable here — treating them as walls made ~60% of stairs look
// unreachable and the AI wandered.

// Monster currently standing on tile (i,j), or null.
function monsterAt(i, j) {
  for (let k = 0; k < state.monsterNumber; k++) {
    const m = state.monsters[k];
    if (m.alive && rint(m.left / SC) === i && rint(m.top / SC) === j) return m;
  }
  return null;
}

// Estimated cost of killing monster m by stepping into it: number of hits and
// the HP we expect to lose to its retaliation over those hits.
// The engine's damage roll is Int(ATK*(0.9..1.1)/DEF)+1 — the +1 and the spread
// both matter. Assume our low roll and their high roll so the estimate errs on
// the safe side (the old estimate dropped the +1, which under-counted incoming
// damage by >=1 per hit and made "safe" fights lethal on deep floors).
function fightCost(m) {
  const { player } = state;
  const myDmg = rint(player.atk * 0.9 / Math.max(1, m.def)) + 1;
  const hits = Math.max(1, Math.ceil(m.hp / myDmg));
  const theirDmg = rint(m.atk * 1.1 / Math.max(1, player.def)) + 1;
  return { hits, hpCost: hits * theirDmg };
}

// Tiles holding monsters too expensive to bump into (losing >=50% of current
// HP). Farming routes treat these as walls so a box run doesn't walk through a
// lethal monster. Under Stealth nothing can hit back, so nothing is dangerous.
function dangerMap() {
  const N = LAND_NUMBER;
  const map = new Uint8Array(N * N);
  if (state.player.condition === A.Stealth) return map;
  for (let k = 0; k < state.monsterNumber; k++) {
    const m = state.monsters[k];
    if (!m.alive) continue;
    if (fightCost(m).hpCost >= state.player.hp * 0.5) {
      map[rint(m.top / SC) * N + rint(m.left / SC)] = 1;
    }
  }
  return map;
}

// Cost of moving onto tile (ni,nj). Infinity = impassable. This is the edge
// weight used to build the field; it must depend only on (fairly) static tile
// state so the field is stable from tick to tick.
function enterCost(ni, nj) {
  const { landsquare, player } = state;
  const ls = landsquare[ni][nj];
  const c = ls.condition;
  if (c === T.Room || c === T.Stair) return 1;
  if (c <= T.PurpleBox) { // box: pay the hits to smash it (red boxes are nasty)
    const myDmg = rint(player.atk * 0.9 / Math.max(1, ls.def)) + 1;
    const breakHits = Math.max(1, Math.ceil(ls.hp / myDmg));
    return breakHits + (c === T.RedBox ? 8 : 1);
  }
  if (c === T.Enemy) {
    const m = monsterAt(ni, nj);
    const fc = m ? fightCost(m) : { hits: 1, hpCost: 0 };
    return 1 + fc.hits + 4 * fc.hpCost / Math.max(1, player.hp);
  }
  return Infinity; // Wall / Mine
}

// Dijkstra distance-to-stair field over the weighted grid above.
function costField() {
  const { landsquare } = state;
  const N = LAND_NUMBER;
  const dist = new Float64Array(N * N).fill(Infinity);
  // binary min-heap of indices keyed by dist
  const heap = []; // stores idx; ordered by dist[idx]
  const push = idx => {
    heap.push(idx);
    let c = heap.length - 1;
    while (c > 0) {
      const p = (c - 1) >> 1;
      if (dist[heap[p]] <= dist[heap[c]]) break;
      [heap[p], heap[c]] = [heap[c], heap[p]]; c = p;
    }
  };
  const pop = () => {
    const top = heap[0], last = heap.pop();
    if (heap.length) {
      heap[0] = last;
      let p = 0;
      for (;;) {
        const l = 2 * p + 1, r = l + 1; let s = p;
        if (l < heap.length && dist[heap[l]] < dist[heap[s]]) s = l;
        if (r < heap.length && dist[heap[r]] < dist[heap[s]]) s = r;
        if (s === p) break;
        [heap[s], heap[p]] = [heap[p], heap[s]]; p = s;
      }
    }
    return top;
  };

  for (let i = 0; i < N; i++)
    for (let j = 0; j < N; j++)
      if (landsquare[i][j].condition === T.Stair) {
        const idx = j * N + i; dist[idx] = 0; push(idx);
      }

  while (heap.length) {
    const cur = pop();
    const ci = cur % N, cj = (cur - ci) / N;
    const cd = dist[cur];
    for (const [di, dj] of DIRS) {
      const ni = ci + di, nj = cj + dj;
      if (ni < 0 || ni >= N || nj < 0 || nj >= N) continue;
      const w = enterCost(ni, nj);
      if (w === Infinity) continue;
      const idx = nj * N + ni;
      const nd = cd + w;
      if (nd < dist[idx]) { dist[idx] = nd; push(idx); }
    }
  }
  return dist;
}

// Step toward the stair by descending the cost field. Because the field is a
// true potential there is always a strictly-lower neighbour leading to the
// stair, so this never oscillates. prevI/prevJ only breaks exact ties.
// Returns 1 (no move) if the stair is unreachable from here.
function navDir(pi, pj, prevI, prevJ) {
  const { landsquare } = state;
  const N = LAND_NUMBER;

  // adjacent stair → step straight onto it
  for (const [di, dj, d] of DIRS) {
    const ni = pi + di, nj = pj + dj;
    if (ni < 0 || ni >= N || nj < 0 || nj >= N) continue;
    if (landsquare[ni][nj].condition === T.Stair) return d;
  }

  const field = costField();
  if (!isFinite(field[pj * N + pi])) return 1; // stair unreachable

  let bestDir = 1, bestCost = Infinity;
  for (const [di, dj, d] of DIRS) {
    const ni = pi + di, nj = pj + dj;
    if (ni < 0 || ni >= N || nj < 0 || nj >= N) continue;
    let f = field[nj * N + ni];
    if (!isFinite(f)) continue;
    if (ni === prevI && nj === prevJ) f += 0.25; // tie-break only
    if (f < bestCost) { bestCost = f; bestDir = d; }
  }
  return bestDir;
}

// When the stair can't be reached even by smashing through boxes, the floor is
// walled off. Breaking boxes is the only way out: box effects can demolish
// walls (re-connecting the stair), hand out a "next floor" charge (PurpleBox),
// or transform the terrain. So head to the nearest box and break it. Returns a
// step toward the nearest reachable box, or 1 if none are reachable.
function navToBoxDir(pi, pj, prevI, prevJ, includeRed) {
  const { landsquare } = state;
  const N = LAND_NUMBER;
  const dist = new Int32Array(N * N).fill(-1);
  const queue = [];
  // seed from box tiles (the targets)
  for (let i = 0; i < N; i++)
    for (let j = 0; j < N; j++) {
      const c = landsquare[i][j].condition;
      if (c <= T.PurpleBox && (includeRed || c !== T.RedBox)) {
        dist[j * N + i] = 0; queue.push(i, j);
      }
    }
  // expand outward over walkable tiles
  let head = 0;
  while (head < queue.length) {
    const ci = queue[head++], cj = queue[head++];
    const cd = dist[cj * N + ci] + 1;
    for (const [di, dj] of DIRS) {
      const ni = ci + di, nj = cj + dj;
      if (ni < 0 || ni >= N || nj < 0 || nj >= N) continue;
      const idx = nj * N + ni;
      if (dist[idx] !== -1) continue;
      const c = landsquare[ni][nj].condition;
      if (c === T.Room || c === T.Stair || c === T.Enemy) { dist[idx] = cd; queue.push(ni, nj); }
    }
  }
  if (dist[pj * N + pi] < 0) return 1; // no box reachable

  let bestDir = 1, bestCost = Infinity;
  for (const [di, dj, d] of DIRS) {
    const ni = pi + di, nj = pj + dj;
    if (ni < 0 || ni >= N || nj < 0 || nj >= N) continue;
    const c = landsquare[ni][nj].condition;
    let fv;
    if (c <= T.PurpleBox && (includeRed || c !== T.RedBox)) fv = 0; // step into box
    else if (c === T.Room || c === T.Stair || c === T.Enemy) {
      fv = dist[nj * N + ni];
      if (fv < 0) continue;
    } else continue;
    if (ni === prevI && nj === prevJ) fv += 0.5;
    if (fv < bestCost) { bestCost = fv; bestDir = d; }
  }
  return bestDir;
}

// ── Farming navigation ───────────────────────────────────────────────────────
// Reaching deep floors requires GRINDING, not rushing the stair. Monster attack
// grows x1.2 per floor (compounding) while the player's defence grows only +1
// per level (linear), so a player that descends at the monster-kills-it-happens-
// to-bump-into rate is chronically under-levelled and dies to deep-floor combat.
// Two facts make farming strongly +EV and cheap:
//   • A floor's monsters do NOT get stronger while you stay on it (statusCheck
//     (true) runs only on floor ENTRY), so clearing the current floor is exp at
//     a fixed difficulty that makes every later floor safer.
//   • Breaking a box grants +10 turns and the turn budget carries across floors,
//     so collecting boxes (esp. GREEN = permanent stat ups, PURPLE = ability
//     charges) pays for itself in turns and powers the player up.
// So before descending we sweep the floor: grab beneficial boxes, then grind
// every monster we can kill safely. We bail to the stair when hurt-with-no-heal
// or low on turns.

// Generic BFS to the nearest tile satisfying isTarget. Seeds from all target
// tiles and floods outward over passable terrain (Room/Stair/Enemy plus boxes,
// which the player smashes through), so dist at the player tile is the path
// length; we then step to the neighbour with the lowest dist. Returns 1 if no
// target is reachable.
function navToTilesDir(pi, pj, isTarget, allowRed, prevI = -1, prevJ = -1, danger = null) {
  const { landsquare } = state;
  const N = LAND_NUMBER;
  const passable = (i, j) => {
    const c = landsquare[i][j].condition;
    if (c === T.Enemy) return !danger || !danger[j * N + i]; // don't route through lethal monsters
    if (c === T.Room || c === T.Stair) return true;
    if (c <= T.PurpleBox) return allowRed || c !== T.RedBox; // breakable box
    return false; // wall / mine
  };
  const dist = new Int32Array(N * N).fill(-1);
  const queue = [];
  for (let i = 0; i < N; i++)
    for (let j = 0; j < N; j++)
      if (isTarget(i, j)) { dist[j * N + i] = 0; queue.push(i, j); }
  let head = 0;
  while (head < queue.length) {
    const ci = queue[head++], cj = queue[head++];
    const cd = dist[cj * N + ci] + 1;
    for (const [di, dj] of DIRS) {
      const ni = ci + di, nj = cj + dj;
      if (ni < 0 || ni >= N || nj < 0 || nj >= N) continue;
      const idx = nj * N + ni;
      if (dist[idx] !== -1 || !passable(ni, nj)) continue;
      dist[idx] = cd; queue.push(ni, nj);
    }
  }
  if (dist[pj * N + pi] < 0) return 1; // no target reachable

  let bestDir = 1, bestCost = Infinity;
  for (const [di, dj, d] of DIRS) {
    const ni = pi + di, nj = pj + dj;
    if (ni < 0 || ni >= N || nj < 0 || nj >= N) continue;
    let fv;
    if (isTarget(ni, nj)) fv = 0;            // step straight onto the target
    else if (passable(ni, nj)) { fv = dist[nj * N + ni]; if (fv < 0) continue; }
    else continue;
    if (ni === prevI && nj === prevJ) fv += 0.5; // damp a->b->a flicker
    if (fv < bestCost) { bestCost = fv; bestDir = d; }
  }
  return bestDir;
}

// Boxes worth detouring for: PURPLE (ability charges), GREEN (mostly permanent
// boosts), and BLUE only while hurt (small heal + cures status). Red is almost
// all downside — skipped. YELLOW measured -EV (seeded A/B, 150 eps: deaths
// 96→107, median 31→25): the atk/2 outcome doubles the retaliation taken in
// every later fight and wallbreak/boxattack monsters wreck the floor, which
// together outweigh the rare atk/def doubles. Don't re-add without re-measuring.
function beneficialBoxAt(i, j) {
  const { player } = state;
  const c = state.landsquare[i][j].condition;
  if (c === T.PurpleBox) return true;
  if (c === T.GreenBox && state.floor < 300) return true;
  if (c === T.BlueBox && player.hp / player.maxHp < 0.85) return true;
  return false;
}

// A monster we can kill while keeping an HP buffer — i.e. affordable exp.
// Death diagnostics showed deaths cluster on floors 11-30 right after the old
// strict bar (hits<=4, cost<40% maxHp) cut off ALL grinding, starving the exp
// flow exactly where monster atk compounding bites. The bar is now adaptive:
// commit to a fight as long as the projected retaliation leaves a 25%-of-maxHp
// buffer from CURRENT hp, with a hit cap so long fights don't let the rest of
// the floor converge on us mid-fight.
function safeMonsterAt(i, j) {
  if (state.landsquare[i][j].condition !== T.Enemy) return false;
  const m = monsterAt(i, j);
  if (!m) return false;
  if (state.player.condition === A.Stealth) return true; // can't be hit back
  const { player } = state;
  const fc = fightCost(m);
  return fc.hits <= 8 && fc.hpCost < player.hp - player.maxHp * 0.25;
}

// ── Monster target lock ──────────────────────────────────────────────────────
// Monsters MOVE, so re-running "BFS to the nearest safe monster" every tick
// makes the nearest target flip between two monsters on opposite sides and the
// player ping-pongs between them (one floor logged 882 reversals in 1082
// ticks). Lock onto one monster and chase IT until it dies, stops being safe,
// or becomes unreachable. Locked targets also chase us, so distance converges.
let _lockIdx = -1;

function lockedMonsterDir(pi, pj, danger) {
  const { monsters } = state;
  if (_lockIdx >= 0) {
    const m = monsters[_lockIdx];
    if (!m.alive || !safeMonsterAt(rint(m.left / SC), rint(m.top / SC))) _lockIdx = -1;
  }
  if (_lockIdx < 0) {
    let best = Infinity;
    for (let k = 0; k < state.monsterNumber; k++) {
      const m = monsters[k];
      if (!m.alive) continue;
      const ti = rint(m.left / SC), tj = rint(m.top / SC);
      if (!safeMonsterAt(ti, tj)) continue;
      const d = Math.abs(ti - pi) + Math.abs(tj - pj);
      if (d < best) { best = d; _lockIdx = k; }
    }
  }
  if (_lockIdx < 0) return 1;
  const m = monsters[_lockIdx];
  const ti = rint(m.left / SC), tj = rint(m.top / SC);
  const dir = navToTilesDir(pi, pj, (i, j) => i === ti && j === tj, false, _prevI, _prevJ, danger);
  if (dir === 1) _lockIdx = -1; // unreachable — drop the lock
  return dir;
}

// Monsters chase at equal speed, so you cannot outrun an adjacent one — fleeing
// only prolongs exposure and walks you away from the stair (a death spiral at
// low HP). The right move is almost always to KILL adjacent monsters: each kill
// removes an attacker and the exp levels you up, which heals. Returns a
// direction toward the easiest adjacent monster we can kill without the
// retaliation outright killing us (under Stealth monsters can't hit back, so
// any kill is free), preferring the fewest hits. 1 = no worthwhile target.
function bestKillDir(pi, pj) {
  const { player } = state;
  const isStealth = player.condition === A.Stealth;
  let bestDir = 1, bestHits = Infinity, bestCost = Infinity;
  for (const [di, dj, d] of DIRS) {
    const ni = pi + di, nj = pj + dj;
    if (ni < 0 || ni >= LAND_NUMBER || nj < 0 || nj >= LAND_NUMBER) continue;
    if (state.landsquare[ni][nj].condition !== T.Enemy) continue;
    const m = monsterAt(ni, nj);
    if (!m) continue;
    const fc = fightCost(m);
    if (!isStealth && fc.hpCost >= player.hp) continue; // would kill us — skip
    if (fc.hits < bestHits || (fc.hits === bestHits && fc.hpCost < bestCost)) {
      bestHits = fc.hits; bestCost = fc.hpCost; bestDir = d;
    }
  }
  return bestDir;
}

// ── Adjacent box pickup ──────────────────────────────────────────────────────

function adjacentBoxPickup(pi, pj) {
  const hpRatio = state.player.hp / state.player.maxHp;
  for (const [di, dj, d] of DIRS) {
    const ni = pi + di, nj = pj + dj;
    if (ni < 0 || ni >= LAND_NUMBER || nj < 0 || nj >= LAND_NUMBER) continue;
    const c = state.landsquare[ni][nj].condition;
    if (c === T.PurpleBox) return d;
    if (c === T.GreenBox && state.floor < 300) return d;
    if (c === T.BlueBox && hpRatio < 0.85) return d;
  }
  return 1;
}

// ── Situation awareness ──────────────────────────────────────────────────────

function nearbyMonsterCount(pi, pj, dist) {
  let count = 0;
  for (let i = 0; i < state.monsterNumber; i++) {
    const m = state.monsters[i];
    if (!m.alive) continue;
    if (Math.abs(rint(m.left / SC) - pi) + Math.abs(rint(m.top / SC) - pj) <= dist) count++;
  }
  return count;
}

function fleeDir(pi, pj) {
  let nearI = -1, nearJ = -1, nearDist = Infinity;
  for (let i = 0; i < state.monsterNumber; i++) {
    const m = state.monsters[i];
    if (!m.alive) continue;
    const mi = rint(m.left / SC), mj = rint(m.top / SC);
    const d = Math.abs(mi - pi) + Math.abs(mj - pj);
    if (d < nearDist) { nearDist = d; nearI = mi; nearJ = mj; }
  }
  if (nearI < 0) return randomDir(pi, pj);

  let bestDir = 1, bestScore = -Infinity;
  for (const [di, dj, d] of DIRS) {
    const ni = pi + di, nj = pj + dj;
    if (ni < 0 || ni >= LAND_NUMBER || nj < 0 || nj >= LAND_NUMBER) continue;
    const c = state.landsquare[ni][nj].condition;
    if (c !== T.Room && c !== T.Stair) continue;
    const score = Math.abs(ni - nearI) + Math.abs(nj - nearJ);
    if (score > bestScore) { bestScore = score; bestDir = d; }
  }
  return bestDir > 1 ? bestDir : randomDir(pi, pj);
}

function randomDir(pi, pj) {
  const candidates = [];
  for (const [di, dj, d] of DIRS) {
    const ni = pi + di, nj = pj + dj;
    if (ni < 0 || ni >= LAND_NUMBER || nj < 0 || nj >= LAND_NUMBER) continue;
    const c = state.landsquare[ni][nj].condition;
    if (c === T.Room || c === T.Enemy || c === T.Stair) candidates.push(d);
  }
  if (candidates.length === 0) return 1;
  return candidates[rint(Math.random() * candidates.length)];
}

// ── Main decision loop ───────────────────────────────────────────────────────
// Priority order (top = highest):
//   1. Survive: HP/turn critical, surrounded
//   2. Tactical opening: convert dense monster floors to boxes
//   3. Escape cycles: ability spend or flee
//   4. Navigate: opportunistic box, then stair

export function autoTick() {
  const { player, abilityHp } = state;

  // Auto-retry on death: linger on the GameOver screen briefly, then press X to
  // return to the title screen. The next tick (now at Entrance) presses Enter to
  // begin a fresh run — the same two keystrokes a player would use.
  if (state.gameMode === GameMode.GameOver) {
    if (++_retryDelay >= RETRY_DELAY_TICKS) { _retryDelay = 0; pressKey('KeyX'); }
    return;
  }
  if (state.gameMode === GameMode.Entrance) {
    pressKey('Enter');
    _lastFloor = -1; // force per-floor bookkeeping reset on the fresh run
    return;
  }
  _retryDelay = 0;

  if (state.gameMode !== GameMode.Dungeon) return;
  if (!player.alive) return;
  if (player.direction > 1) return;

  // New floor: reset bookkeeping
  if (state.floor !== _lastFloor) {
    resetCycleState();
    _lastFloor = state.floor;
    _floorOpenedHandled = false;
  }

  const pi = rint(player.left / SC);
  const pj = rint(player.top / SC);
  updateCycleDetection(progressSig(pi, pj));

  const hpRatio = player.hp / player.maxHp;
  const adjMonsters = nearbyMonsterCount(pi, pj, 1);
  const aliveMonsters = countAliveMonsters();

  // ── Priority 1: SURVIVE ──────────────────────────────────────────────

  if (hpRatio < 0.4 && abilityHp[1] > 0) {
    abilityEffect('KeyX'); resetCycleState(); return;
  }

  if (hpRatio < 0.3 && abilityHp[1] === 0) {
    if (abilityHp[3] > 0 && aliveMonsters > 0) { abilityEffect('KeyD'); resetCycleState(); return; }
    if (abilityHp[4] > 0) { abilityEffect('Enter'); resetCycleState(); return; }
    if (abilityHp[0] > 0 && aliveMonsters > 0 && nearbyMonsterCount(pi, pj, 5) > 0) {
      abilityEffect('KeyZ'); resetCycleState(); return;
    }
  }

  if (state.turn < 50 && abilityHp[4] > 0) {
    abilityEffect('Enter'); resetCycleState(); return;
  }

  if (adjMonsters >= 3 && abilityHp[0] > 0) {
    abilityEffect('KeyZ'); resetCycleState(); return;
  }

  // ── Priority 2: TACTICAL OPENING ─────────────────────────────────────

  if (!_floorOpenedHandled && state.monsterNumber >= 30 && abilityHp[3] > 0) {
    abilityEffect('KeyD');
    _floorOpenedHandled = true;
    resetCycleState();
    return;
  }
  _floorOpenedHandled = true;

  // ── Priority 3: ESCAPE CYCLES ────────────────────────────────────────

  if (_stuckTicks >= STUCK_THRESHOLD) {
    resetCycleState();
    // Only spend monster-targeting abilities if monsters are actually alive
    if (aliveMonsters > 0 && abilityHp[3] > 0) { abilityEffect('KeyD'); return; }
    if (aliveMonsters > 0 && abilityHp[0] > 0) { abilityEffect('KeyZ'); return; }
    if (abilityHp[4] > 0) { abilityEffect('Enter'); return; }
    if (abilityHp[2] > 0) { abilityEffect('KeyC'); return; }
    // genuinely stuck and no abilities: break boxes to try to open a route
    let dir = navToBoxDir(pi, pj, _prevI, _prevJ, false);
    if (dir === 1) dir = navToBoxDir(pi, pj, _prevI, _prevJ, true);
    if (dir === 1) dir = fleeDir(pi, pj);
    if (dir === 1) dir = randomDir(pi, pj);
    if (dir > 1) { player.direction = dir; _prevI = pi; _prevJ = pj; }
    return;
  }

  // ── Priority 4: NAVIGATE ────────────────────────────────────────────

  // Adjacent stair → just step onto it and end the floor.
  for (const [di, dj, d] of DIRS) {
    const ni = pi + di, nj = pj + dj;
    if (ni < 0 || ni >= LAND_NUMBER || nj < 0 || nj >= LAND_NUMBER) continue;
    if (state.landsquare[ni][nj].condition === T.Stair) {
      player.direction = d; _prevI = pi; _prevJ = pj; return;
    }
  }

  // Kill an adjacent monster. Chasers move at the player's speed, so they can't
  // be outrun; killing removes the attacker and the exp levels the player up
  // (which heals). This is what keeps the player alive — fleeing does not.
  const killDir = bestKillDir(pi, pj);
  if (killDir > 1) {
    player.direction = killDir; _prevI = pi; _prevJ = pj; return;
  }

  // Opportunistically grab a beneficial adjacent box.
  const boxDir = adjacentBoxPickup(pi, pj);
  if (boxDir > 1) {
    player.direction = boxDir; _prevI = pi; _prevJ = pj; return;
  }

  // ── FARM vs DESCEND ──────────────────────────────────────────────────
  // Sweep the floor for resources before taking the stair, unless we're hurt
  // with no way to heal or running low on turns — then cut losses and descend.
  const lowTurn = state.turn < 60;
  const hurt = hpRatio < 0.5;
  // KeyC (全消去) clears boxes/walls/monsters but does NOT heal HP (VB6-
  // faithful), so the only real heals are KeyX charges and blue boxes.
  const noHeal = abilityHp[1] === 0;

  let dir = 1;
  if (!lowTurn) {
    const danger = dangerMap();
    if (hurt && noHeal) {
      // Hurt with no heal charges: descending now just spawns a fresh,
      // stronger floor at low HP. Top up on blue boxes first (each heals
      // maxHp/10); descend only when none are reachable.
      dir = navToTilesDir(pi, pj,
        (i, j) => state.landsquare[i][j].condition === T.BlueBox,
        false, _prevI, _prevJ, danger);
    } else {
      // 1) collect a beneficial box (heal/boost + 10 turns), then
      // 2) grind a reachable monster we can kill safely, while HP is healthy.
      dir = navToTilesDir(pi, pj, beneficialBoxAt, false, _prevI, _prevJ, danger);
      if (dir === 1 && hpRatio > 0.5 && aliveMonsters > 0)
        dir = lockedMonsterDir(pi, pj, danger);
    }
  }

  if (dir === 1) {
    // Floor farmed out (or bailing): head for the stair.
    dir = navDir(pi, pj, _prevI, _prevJ);
    if (dir === 1) {
      // stair unreachable from here (walled-off floor) — escape by breaking boxes
      if (abilityHp[4] > 0) { abilityEffect('Enter'); return; }
      dir = navToBoxDir(pi, pj, _prevI, _prevJ, false);
      if (dir === 1) dir = navToBoxDir(pi, pj, _prevI, _prevJ, true); // allow red boxes
      if (dir === 1) dir = randomDir(pi, pj);
    }
  }

  if (dir > 1) player.direction = dir;
  _prevI = pi; _prevJ = pj;
}
