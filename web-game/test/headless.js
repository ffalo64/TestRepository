// Headless auto-play harness. Runs the real engine + autoPlay loop in Node
// (no canvas/audio/images) to measure balance: floor reached, deaths, 1000F
// clear rate.
//
//   node test/headless.js [episodes] [maxTicksPerEpisode]
//   node test/headless.js trace [maxTicks]
//   node test/headless.js mapinfo [samples]
//
// Env: SEED0=<int> offsets the per-episode RNG seeds so parallel workers
// sample disjoint episode sets. JSON=1 prints a single machine-readable line.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { state } from '../src/state.js';
import { GameMode, T, SC, LAND_NUMBER } from '../src/constants.js';
import { landSet, monsterSet, floorSet, movement, statusCheck } from '../src/engine.js';
import { autoTick } from '../src/autoPlay.js';

// Load the real Map40.bmp into state so the harness uses the same maps as the
// browser (the engine otherwise falls back to a different procedural generator).
function loadRealMap(path) {
  let buf;
  try { buf = readFileSync(path); } catch { return false; }
  if (buf[0] !== 0x42 || buf[1] !== 0x4d) return false; // 'BM'
  const off = buf.readUInt32LE(10);
  const w = buf.readInt32LE(18), h = buf.readInt32LE(22), bpp = buf.readUInt16LE(28);
  if (bpp !== 24) return false;
  const rowSize = Math.floor((bpp * w + 31) / 32) * 4; // padded to 4 bytes
  const rgba = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) {
    const srcRow = off + (h - 1 - y) * rowSize; // BMP is bottom-up
    for (let x = 0; x < w; x++) {
      const s = srcRow + x * 3;
      const d = (y * w + x) * 4;
      rgba[d] = buf[s + 2]; rgba[d + 1] = buf[s + 1]; rgba[d + 2] = buf[s]; rgba[d + 3] = 255;
    }
  }
  state.mapPixels = rgba; state.mapW = w; state.mapH = h;
  // colorSet: reference colours live at y=1, x=1..6
  for (let x = 1; x <= 6; x++) {
    const i = (1 * w + x) * 4;
    state.colorRef[x] = { r: rgba[i], g: rgba[i + 1], b: rgba[i + 2] };
  }
  return true;
}

const rint = n => Math.floor(n);
const DIRS = [[0, -1], [0, 1], [1, 0], [-1, 0]];
const N = LAND_NUMBER;

// Deterministic RNG: reseed Math.random per episode so different balance
// variants are measured on the SAME map sequences and monster rolls.
function mulberry32(a) {
  return function () {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// distance field from every stair tile over walkable tiles (Room/Enemy/Stair)
function distField() {
  const dist = new Int32Array(N * N).fill(-1);
  const q = [];
  for (let i = 0; i < N; i++)
    for (let j = 0; j < N; j++)
      if (state.landsquare[i][j].condition === T.Stair) { dist[j * N + i] = 0; q.push(i, j); }
  let h = 0;
  while (h < q.length) {
    const ci = q[h++], cj = q[h++], cd = dist[cj * N + ci] + 1;
    for (const [di, dj] of DIRS) {
      const ni = ci + di, nj = cj + dj;
      if (ni < 0 || ni >= N || nj < 0 || nj >= N) continue;
      const idx = nj * N + ni;
      if (dist[idx] !== -1) continue;
      const c = state.landsquare[ni][nj].condition;
      if (c === T.Room || c === T.Enemy || c === T.Stair) { dist[idx] = cd; q.push(ni, nj); }
    }
  }
  return dist;
}

function playerTile() {
  return [rint(state.player.left / SC), rint(state.player.top / SC)];
}

function startRun() {
  // full reset: Entrance landSet zeroes floor / abilityHp / player stats,
  // otherwise the module-singleton state leaks between episodes.
  state.gameMode = GameMode.Entrance;
  landSet();
  state.gameMode = GameMode.Dungeon;
  landSet(); monsterSet(); floorSet();
  state.autoPlay = true;
}

function monsterHpSum() {
  let s = 0;
  for (let i = 0; i < state.monsterNumber; i++)
    if (state.monsters[i].alive) s += state.monsters[i].hp;
  return s;
}

function runEpisode(maxTicks) {
  startRun();
  let curFloor = state.floor;
  const death = { cause: '-', floor: 0, turn: 0 };
  // instrumentation: revive-economy tracking
  let reviveSpent = 0, reviveGained = 0, hpOneEvents = 0;
  let prevRevive = state.abilityHp[5];
  // turn-economy tracking (deep floors >= 300)
  let deepMoves = 0, deepBoxGains = 0, deepFloors = 0, lastDeepFloor = -1;
  let prevTurn = state.turn;

  for (let tick = 0; tick < maxTicks; tick++) {
    if (state.gameMode !== GameMode.Dungeon) break; // GameOver / GameClear
    if (state.floor !== curFloor) curFloor = state.floor;

    const turnBefore = state.turn;
    autoTick();
    movement();
    statusCheck(false);

    const nowRevive = state.abilityHp[5];
    if (nowRevive < prevRevive) reviveSpent += prevRevive - nowRevive;
    else if (nowRevive > prevRevive) reviveGained += nowRevive - prevRevive;
    prevRevive = nowRevive;
    if (state.player.hp === 1) hpOneEvents++;

    if (state.floor >= 300) {
      if (state.floor !== lastDeepFloor) { deepFloors++; lastDeepFloor = state.floor; }
      const dTurn = state.turn - prevTurn;
      // one tick can combine a -1 move with +15s from boxes; decompose
      if (dTurn < 0) deepMoves += -dTurn;
      else if (dTurn > 0) deepBoxGains += dTurn;
    }
    prevTurn = state.turn;

    if (state.gameMode === GameMode.GameOver && death.cause === '-') {
      death.cause = state.turn <= 0 ? 'turn' : 'hp';
      death.floor = curFloor;
      death.lvl = state.player.level;
      death.def = state.player.def;
      death.atk = state.player.atk;
      death.maxHp = state.player.maxHp;
      death.mAtk = state.monsters[0].atk;
      death.mHp = state.monsters[0].maxHp;
      death.mNum = state.monsterNumber;
      death.ab = state.abilityHp.slice(0, 6);
    }
  }

  const outcome = state.gameMode === GameMode.GameClear ? 'CLEAR'
    : state.gameMode === GameMode.GameOver ? 'DEAD' : 'TIMEOUT';
  death.reviveSpent = reviveSpent;
  death.reviveGained = reviveGained;
  death.hpOneEvents = hpOneEvents;
  death.deepMovesPerFloor = deepFloors ? +(deepMoves / deepFloors).toFixed(1) : 0;
  death.deepGainPerFloor = deepFloors ? +(deepBoxGains / deepFloors).toFixed(1) : 0;
  return { outcome, deepestFloor: curFloor, death };
}

// ── trace mode ──────────────────────────────────────────────────────────────
function trace(maxTicks) {
  startRun();
  let curFloor = state.floor;
  let floorEntry = { hp: state.player.hp, lvl: state.player.level, turn: state.turn };
  for (let tick = 0; tick < maxTicks; tick++) {
    if (state.gameMode !== GameMode.Dungeon) break;
    if (state.floor !== curFloor) {
      console.log(`floor ${curFloor}: enter L${floorEntry.lvl} hp${rint(floorEntry.hp)} -> exit L${state.player.level} hp${rint(state.player.hp)}/${state.player.maxHp} turn${state.turn} mNum=${state.monsterNumber} mAtk=${state.monsters[0].atk} ab=[${state.abilityHp.slice(0,6)}]`);
      curFloor = state.floor;
      floorEntry = { hp: state.player.hp, lvl: state.player.level, turn: state.turn };
    }
    autoTick(); movement(); statusCheck(false);
  }
  const why = state.gameMode === GameMode.GameOver ? 'GAME OVER' : state.gameMode === GameMode.GameClear ? 'CLEAR' : 'TIMEOUT';
  console.log(`\n>>> ${why} at floor ${curFloor}, turn ${state.turn}, L${state.player.level} atk=${state.player.atk} def=${state.player.def} maxHp=${state.player.maxHp}`);
}

// ── main ──────────────────────────────────────────────────────────────────────
const mapPath = fileURLToPath(new URL('../public/images/Map40.bmp', import.meta.url));
const realMap = loadRealMap(mapPath);
if (!process.env.JSON) console.log(`map source: ${realMap ? 'Map40.bmp (real)' : 'procedural fallback'}`);

if (process.argv[2] === 'trace') { trace(parseInt(process.argv[3] || '2000', 10)); process.exit(0); }

const episodes = parseInt(process.argv[2] || '20', 10);
const maxTicks = parseInt(process.argv[3] || '400000', 10);
const seed0 = parseInt(process.env.SEED0 || '0', 10);

const outcomes = { CLEAR: 0, DEAD: 0, TIMEOUT: 0 };
const deepest = [];
const deaths = [];

for (let e = 0; e < episodes; e++) {
  Math.random = mulberry32(0x9e3779b9 + (seed0 + e) * 7919);
  const r = runEpisode(maxTicks);
  outcomes[r.outcome]++;
  deepest.push(r.deepestFloor);
  if (r.outcome === 'DEAD') deaths.push(r.death);
  if (!process.env.JSON) console.log(`  ep${seed0 + e}: ${r.outcome} floor=${r.deepestFloor}`);
}

if (process.env.JSON) {
  console.log(JSON.stringify({ seed0, episodes, outcomes, deepest, deaths }));
  process.exit(0);
}

deepest.sort((a, b) => a - b);
const median = deepest[rint(deepest.length / 2)];
console.log('═══ Balance report ═══');
console.log(`episodes=${episodes}  outcomes: CLEAR=${outcomes.CLEAR} DEAD=${outcomes.DEAD} TIMEOUT=${outcomes.TIMEOUT}`);
console.log(`1000F reach rate: ${(100 * outcomes.CLEAR / episodes).toFixed(0)}%`);
console.log(`floor reached: median=${median} min=${deepest[0]} max=${deepest[deepest.length - 1]}`);
if (deaths.length) {
  const avg = f => (deaths.reduce((a, d) => a + f(d), 0) / deaths.length).toFixed(1);
  const floors = deaths.map(d => d.floor).sort((a, b) => a - b);
  const q = p => floors[rint(floors.length * p)];
  console.log(`death cause: hp=${deaths.filter(d => d.cause === 'hp').length} turn=${deaths.filter(d => d.cause === 'turn').length}`);
  console.log(`death floor quartiles: p25=${q(0.25)} p50=${q(0.5)} p75=${q(0.75)}`);
  console.log(`at death: lvl=${avg(d => d.lvl)} atk=${avg(d => d.atk)} def=${avg(d => d.def)} maxHp=${avg(d => d.maxHp)}`);
  console.log(`monsters: atk=${avg(d => d.mAtk)} maxHp=${avg(d => d.mHp)} num=${avg(d => d.mNum)}`);
  console.log(`unspent at death: [${['全体','回復','全回','除去','次階','蘇生'].map((n, k) => `${n}=${avg(d => d.ab[k])}`).join(' ')}]`);
}
