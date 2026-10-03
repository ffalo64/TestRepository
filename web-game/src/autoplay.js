// オートプレイ（ブラウザ版の追加機能。VB6 版には無い）。
// 毎フレーム tick() の前に nextKeys(g) を呼び、返ってきたキーを keyDown() に渡す。
// DOM 非依存なので test/autoplay.js からも使える。
//
// 方針
// - 戦略: 箱ごとに「今の状態で開けたときの期待値（ターン換算）」を出し、
//   Dijkstra の手数との差が一番得な箱か階段を目指す。
// - 戦術: モンスターの動き（MovementSub / PositionCheck）は決定的なので、
//   近くのモンスターの動きを数手先までシミュレーションして、被弾しない手を選ぶ。
// - コマンド: 避けきれないときは 次の階へ / 箱化 / 全消去 / 全快 で切り抜ける。

import {
  ENTRANCE, DUNGEON, HOWTO_PLAY, OPTIONS, GAME_OVER, GAME_CLEAR, MUSEUM, BLESSING,
  BLUE_BOX, RED_BOX, YELLOW_BOX, GREEN_BOX, PURPLE_BOX, STAIR, WALL, ROOM, ENEMY,
  LAND_NUMBER, WALL_BREAK, SLOW, STEALTH, TURN_CONST,
} from './constants.js';

const N = LAND_NUMBER;
const MENU_WAIT = 30; // メニュー画面で次のキーを押すまでのフレーム数（1.5 秒）
const RESULT_WAIT = 80; // ゲームオーバー/クリア画面を見せるフレーム数（4 秒）

// 8 方向と、それを入力するキー（engine の Direction と同じ素数の積）
const STEPS = [
  [0, -1, ['Up'], 2], [0, 1, ['Down'], 3], [1, 0, ['Right'], 5], [-1, 0, ['Left'], 7],
  [1, -1, ['Up', 'Right'], 10], [1, 1, ['Down', 'Right'], 15],
  [-1, 1, ['Down', 'Left'], 21], [-1, -1, ['Up', 'Left'], 14],
];
// engine.js の DETOURS と同じ（進めなかったモンスターの回り込み先）
const DETOURS = {
  2: [[-1, -1], [1, -1]], 3: [[1, 1], [-1, 1]], 5: [[1, -1], [1, 1]], 7: [[-1, 1], [-1, -1]],
  10: [[0, -1], [1, 0]], 15: [[1, 0], [0, 1]], 21: [[0, 1], [-1, 0]], 14: [[-1, 0], [0, -1]],
};

// コマンド 1 回分の価値（ターン換算）。[全体攻撃, Hp全快, 全消去, 箱化, 次の階へ, 復活の珠]
const CMD_VALUE = [10, 25, 50, 30, 60, 120];
const SEARCH_RADIUS = 9; // これより遠いモンスターは先読みでは動かないものとして扱う
const DEATH = 1e5;
const STALL_SOFT = 6; // これだけ進展が無ければ「敵を待って倒す」のをやめる
const STALL_HARD = 20; // これだけ進展が無ければコマンドで打開する

const isEdge = (x, y) => x === 0 || y === 0 || x === N - 1 || y === N - 1;
const cheb = (ax, ay, bx, by) => Math.max(Math.abs(ax - bx), Math.abs(ay - by));

// calcDamage の期待値と最大値
const avgDamage = (atk, def) => Math.floor(atk / Math.max(1, def)) + 1;
const maxDamage = (atk, def) => Math.floor((atk * 1.1) / Math.max(1, def)) + 1;
const minDamage = (atk, def) => Math.floor((atk * 0.9) / Math.max(1, def)) + 1;
const hitsToBreak = (atk, t) => Math.ceil(Math.max(1, t.hp) / avgDamage(atk, t.def));

// 最小ヒープ付き Dijkstra。cost[i] はマス i に「入る」コスト（箱は壊す手数込み、壁は Infinity）。
// stop[i] のマス（階段）からは先へ広げない。
// reverse=true なら src に向かう距離（各マスから src まで）を求める。
function dijkstra(cost, stop, src, reverse) {
  const dist = new Float64Array(N * N).fill(Infinity);
  const prev = new Int32Array(N * N).fill(-1);
  const hd = [];
  const hi = [];
  const push = (d, i) => {
    let k = hd.length;
    hd.push(d);
    hi.push(i);
    while (k > 0) {
      const pk = (k - 1) >> 1;
      if (hd[pk] <= d) break;
      hd[k] = hd[pk]; hi[k] = hi[pk];
      k = pk;
    }
    hd[k] = d; hi[k] = i;
  };
  const pop = () => {
    const d0 = hd[0];
    const i0 = hi[0];
    const d = hd.pop();
    const i = hi.pop();
    if (hd.length) {
      let k = 0;
      for (;;) {
        const l = k * 2 + 1;
        if (l >= hd.length) break;
        const r = l + 1;
        const m = r < hd.length && hd[r] < hd[l] ? r : l;
        if (hd[m] >= d) break;
        hd[k] = hd[m]; hi[k] = hi[m];
        k = m;
      }
      hd[k] = d; hi[k] = i;
    }
    return [d0, i0];
  };
  dist[src] = 0;
  push(0, src);
  while (hd.length) {
    const [d, i] = pop();
    if (d > dist[i]) continue;
    if (i !== src && stop[i]) continue;
    const x = (i / N) | 0;
    const y = i % N;
    for (const [dx, dy] of STEPS) {
      const nx = x + dx;
      const ny = y + dy;
      if (nx < 0 || ny < 0 || nx >= N || ny >= N) continue;
      const j = nx * N + ny;
      // 前向き: j に入るコスト。逆向き: j から i に入るコスト
      const c = reverse ? (cost[j] === Infinity || stop[j] ? Infinity : cost[i]) : cost[j];
      if (c === Infinity) continue;
      if (d + c < dist[j]) {
        dist[j] = d + c;
        prev[j] = i;
        push(d + c, j);
      }
    }
  }
  return { dist, prev };
}

export function createAutoPlayer() {
  let wait = 0;
  let lastMode = -1;
  let lastFloor = -1;
  // 足踏みの検出用（階ごとにリセット）
  let stall = 0;
  let prevTurn = 0;
  let prevAlive = 0;
  let prevTarget = -1;
  let prevTargetHp = 0;
  const bestDist = new Map(); // 目的地ごとの、これまでで一番近づいた距離

  // ---- 箱の価値（ターン換算）
  function turnWorth(g) {
    if (g.turn < 150) return 2;
    if (g.turn < 400) return 1;
    if (g.turn < 1500) return 0.5;
    if (g.turn < 5000) return 0.2;
    return 0.05;
  }

  // 1 手使うことの重さ。ターンが余っているほど軽い（= 遠くの箱も取りに行く）
  function stepWeight(g) {
    return Math.max(0.2, Math.min(1.5, 300 / Math.max(1, g.turn)));
  }

  function cmdSum(g, f) {
    let s = 0;
    for (let k = 0; k < 6; k++) s += f(g.abilityHp[k], k) * CMD_VALUE[k];
    return s;
  }

  function boxValues(g) {
    const p = g.player;
    const tw = turnWorth(g);
    const gain = p.condition === TURN_CONST ? 0 : 10 * tw;
    const hpLost = 1 - p.hp / Math.max(1, p.maxHp);
    // 紫箱: 出る個数の平均 × 1 個あたりの期待値（全体攻撃 3/15, 全快 4/15, 全消去 1/15, 箱化 1/15, 次の階へ 3/15, 復活 3/15）
    const perCmd = [3, 4, 1, 1, 3, 3].reduce((s, w, k) => s + w * CMD_VALUE[k], 0) / 15;
    const purple = perCmd * (6 - p.ability) / 2;
    // 黄箱 15 通りの期待値
    const yellow = (
      cmdSum(g, () => 1) // 1: 全コマンド +1
      + cmdSum(g, (v) => 5 - v) // 9: 全コマンドが 5 に
      + (1000 - g.turn) * tw // 6: 残りターンが 1000 に
      + 80 + 50 - 50 // 4: 守備×2, 5: 攻撃×2, 7: 攻撃÷2
      + 10 - 20 - 20 - 30 - 5 // 3: 全滅, 2: ターン固定, 8: 壁掘り, 10: 箱消滅, 14: 箱攻撃
    ) / 15;
    const green = (hpLost * 40 + 15 + 10 + 10 + 10 + 15 + 5 + 5 + 3 + 3 + 5 + 3 + 5 + 10) / 15 + 5;
    return {
      turnGain: gain,
      [BLUE_BOX]: 2 + hpLost * 30 + gain,
      // 赤箱は取り返しのつかない効果が多い。ターンが尽きそうなときだけ妥協する
      [RED_BOX]: -60 - (p.hp / Math.max(1, p.maxHp)) * 10 + (g.turn < 60 ? 60 : 0) + gain,
      [YELLOW_BOX]: yellow + gain,
      [GREEN_BOX]: green + gain,
      [PURPLE_BOX]: purple + gain,
    };
  }

  // ---- 先読み用のモンスター移動シミュレーション（engine.js の movement / positionCheckMonster と同じ規則）
  function makeSim(g, local, cond, target, targetHits) {
    const p = g.player;
    const stealth = p.condition === STEALTH;
    const occupied = (mons, x, y, self) => {
      for (let k = 0; k < mons.length; k++) {
        const m = mons[k];
        if (k !== self && m.alive && m.x === x && m.y === y) return true;
      }
      return false;
    };
    const isRoom = (mons, x, y, self) => cond[x * N + y] === ROOM && !occupied(mons, x, y, self);
    const pAvg = (m) => avgDamage(p.atk, m.def);

    // st: { x, y, turn, mons, dmg, dmgMax, hits, kills, done }
    return function step(st, s) {
      const [, , , dir] = STEPS[s];
      let nx = st.x;
      let ny = st.y;
      if (dir % 2 === 0 && ny > 0) ny--;
      if (dir % 3 === 0 && ny < N - 1) ny++;
      if (dir % 5 === 0 && nx < N - 1) nx++;
      if (dir % 7 === 0 && nx > 0) nx--;
      const mons = st.mons.map((m) => ({ ...m }));
      const ns = { ...st, mons, x: st.x, y: st.y, turn: st.turn - 1, k: st.k + 1 };
      const i = nx * N + ny;
      const c = cond[i];
      const mk = mons.findIndex((m) => m.alive && m.x === nx && m.y === ny);
      if (mk >= 0) {
        // 攻撃。倒しきれなければその場に残る
        const m = mons[mk];
        // 安全に倒しきれる相手なら、削った分も評価する（読みの深さより長い戦闘を続けられるように）
        if (m.safe) ns.kills += m.exp * Math.min(1, pAvg(m) / m.hp0);
        m.hp -= pAvg(m);
        // 早く倒すほど高く評価する（「待っていれば倒せる」で足踏みしないように）
        if (m.hp <= 0) { m.alive = false; if (!m.safe) ns.kills += m.exp * (1 - 0.15 * st.k); }
      } else if (c === STAIR) {
        if (i !== target) return null; // 目的地でない階段は踏まない
        ns.done = true;
        return ns;
      } else if (c <= PURPLE_BOX || c === WALL) {
        if (i === target) {
          ns.hits++;
          if (ns.hits >= targetHits) { ns.done = true; return ns; }
        }
      } else if (c === ROOM) {
        ns.x = nx;
        ns.y = ny;
      }
      if (ns.x === st.x && ns.y === st.y && ns.hits === st.hits && mk < 0) ns.waits++;
      // ENEMY（先読み対象外の遠いモンスター）・壁はぶつかるだけ
      if (stealth) return ns;
      for (let k = 0; k < mons.length; k++) {
        const m = mons[k];
        if (!m.alive) continue;
        if (m.ability === SLOW && ns.turn % 2 !== 0) continue;
        const ox = m.x;
        const oy = m.y;
        let d = 1;
        if (m.y > ns.y) { m.y--; d *= 2; }
        if (m.y < ns.y) { m.y++; d *= 3; }
        if (m.x < ns.x) { m.x++; d *= 5; }
        if (m.x > ns.x) { m.x--; d *= 7; }
        if (m.x === ns.x && m.y === ns.y) {
          m.x = ox; m.y = oy;
          ns.dmg += m.avg;
          ns.dmgMax += m.max;
        } else if (!isRoom(mons, m.x, m.y, k) || (m.x === ox && m.y === oy)) {
          const det = DETOURS[d];
          if (det) {
            m.x = ox; m.y = oy;
            for (const [dx, dy] of det) {
              const tx = ox + dx;
              const ty = oy + dy;
              if (tx >= 0 && ty >= 0 && tx < N && ty < N && isRoom(mons, tx, ty, k)) {
                m.x = tx; m.y = ty;
                break;
              }
            }
          }
        }
      }
      return ns;
    };
  }

  // depth 手先まで全探索して、一番良い最初の一手を返す
  function tactical(g, ctx) {
    const { field, target, targetHits, local, cond, hpBudget, deathCost } = ctx;
    const step = makeSim(g, local, cond, target, targetHits);
    const near = local.filter((m) => cheb(m.x, m.y, g.player.x, g.player.y) <= 4).length;
    const depth = local.length === 0 ? 1 : near > 6 ? 3 : 4;
    const dmgW = ctx.dmgScale * 25 / Math.max(1, hpBudget);

    const evalLeaf = (st, k) => {
      if (st.dmgMax >= hpBudget) return deathCost + (depth - k) * -10;
      let sc = st.dmg * dmgW - st.kills * ctx.killW + st.waits * ctx.waitW;
      if (st.done) return sc - 1000 + k;
      sc += field[st.x * N + st.y] - st.hits;
      return sc;
    };
    const rec = (st, k) => {
      if (st.done || k === depth || st.dmgMax >= hpBudget) return evalLeaf(st, k);
      let best = Infinity;
      for (let s = 0; s < 8; s++) {
        const ns = step(st, s);
        if (!ns) continue;
        const v = rec(ns, k + 1);
        if (v < best) best = v;
      }
      return best === Infinity ? evalLeaf(st, k) : best;
    };
    const root = {
      x: g.player.x, y: g.player.y, turn: g.turn, mons: local,
      dmg: 0, dmgMax: 0, hits: 0, kills: 0, waits: 0, k: 0, done: false,
    };
    let best = null;
    for (let s = 0; s < 8; s++) {
      const ns = step(root, s);
      if (!ns) continue;
      const v = rec(ns, 1);
      if (!best || v < best.v) best = { s, v };
    }
    return best;
  }

  function dungeonKeys(g) {
    const p = g.player;
    const ah = g.abilityHp;
    if (g.floor !== lastFloor) {
      lastFloor = g.floor;
      stall = 0;
      prevTurn = g.turn;
      prevAlive = Infinity;
      prevTarget = -1;
      bestDist.clear();
    }

    const living = [];
    for (let i = 0; i < g.monsterNumber; i++) if (g.monsters[i].alive) living.push(g.monsters[i]);
    const stealth = p.condition === STEALTH;
    const lethal = (m) => maxDamage(m.atk, p.def) >= p.hp;

    // ---- 地形とコスト
    const values = boxValues(g);
    const cond = new Int8Array(N * N);
    const cost = new Float64Array(N * N);
    const stop = new Uint8Array(N * N);
    let stair = -1;
    let worth = 0; // 壊す手間より価値が大きい箱の数
    for (let x = 0; x < N; x++) {
      for (let y = 0; y < N; y++) {
        const t = g.land[x][y];
        const i = x * N + y;
        const c = t.condition;
        cond[i] = c;
        if (c === STAIR) { cost[i] = 1; stop[i] = 1; stair = i; }
        else if (c <= PURPLE_BOX) {
          // 箱は壊して進める。損な箱（赤箱など）は通り抜けるだけでも価値の分だけ重くする
          const hits = hitsToBreak(p.atk, t);
          cost[i] = hits + 1 + Math.max(0, -values[c]);
          if (values[c] > hits) worth++;
        } else if (c === WALL) cost[i] = p.condition === WALL_BREAK && !isEdge(x, y) ? hitsToBreak(p.atk, t) + 1 : Infinity;
        else cost[i] = 1;
      }
    }
    const plain = cost.slice(); // モンスターを考えないコスト
    // モンスターの周りを通りにくくする（倒せる相手は軽く、即死級は重く）
    if (!stealth) {
      for (const m of living) {
        if (cheb(m.x, m.y, p.x, p.y) > 20) continue;
        const kill = Math.ceil(m.hp / avgDamage(p.atk, m.def));
        const dmg = avgDamage(m.atk, p.def);
        const danger = lethal(m) ? 30 : Math.min(30, ((kill - 1) * dmg / Math.max(1, p.hp)) * 20);
        for (let dx = -1; dx <= 1; dx++) {
          for (let dy = -1; dy <= 1; dy++) {
            const x = m.x + dx;
            const y = m.y + dy;
            if (x < 0 || y < 0 || x >= N || y >= N) continue;
            const i = x * N + y;
            if (cost[i] !== Infinity) cost[i] += danger;
          }
        }
        cost[m.x * N + m.y] += kill + danger;
      }
    }

    const start = p.x * N + p.y;
    const fwd = dijkstra(cost, stop, start, false);
    const stairDist = stair >= 0 ? fwd.dist[stair] : Infinity;

    // ---- 目的地選び: 箱は「寄り道の手数（箱経由で階段へ − 直接階段へ）− 価値」が一番小さいもの
    const toStair = stair >= 0 ? dijkstra(cost, stop, stair, true).dist : null;
    const base = stairDist < Infinity ? stairDist : 0;
    let goal = -1;
    let bestScore = Infinity;
    if (stair >= 0 && stairDist < g.turn) {
      goal = stair;
      bestScore = 0;
    }
    const wallBreak = p.condition === WALL_BREAK;
    const sw = stepWeight(g);
    for (let i = 0; i < N * N; i++) {
      const c = cond[i];
      if (c > PURPLE_BOX && !(wallBreak && c === WALL)) continue;
      if (fwd.dist[i] >= g.turn) continue; // ターン切れ前に壊せる箱だけ
      const after = toStair && toStair[i] < Infinity ? toStair[i] : base;
      // 壁掘り中は壁も 1 枚 +10 ターン
      const value = c === WALL ? values.turnGain : values[c];
      const score = (fwd.dist[i] + after - base) * sw - value;
      if (score < bestScore) { bestScore = score; goal = i; }
    }

    // 近くの（先読みで動かす）モンスター
    const local = [];
    if (!stealth) {
      for (let i = 0; i < g.monsterNumber; i++) {
        const m = g.monsters[i];
        if (!m.alive || cheb(m.x, m.y, p.x, p.y) > SEARCH_RADIUS) continue;
        const avg = avgDamage(m.atk, p.def);
        const hitsLeft = Math.ceil(m.hp / avgDamage(p.atk, m.def));
        local.push({
          x: m.x, y: m.y, hp: m.hp, hp0: m.maxHp, def: m.def, exp: m.exp, alive: true, ability: m.ability,
          avg, max: maxDamage(m.atk, p.def),
          safe: hitsLeft > 1 && hitsLeft * avg < p.hp * 0.4,
        });
        cond[m.x * N + m.y] = ROOM; // 先読みではモンスター側の座標で管理する
      }
    }

    // 行き先が無い（階段に届かずターンも足りない）→ 次の階へ / 全消去
    if (goal < 0) {
      if (ah[4] > 0) return ['Enter'];
      if (ah[2] > 0) return ['C'];
      return STEPS[(g.turn * 7) % 8][2];
    }

    // 途中に箱があるなら、まず手前の箱が当面の目標
    let target = goal;
    for (let j = goal; j !== start && j >= 0; j = fwd.prev[j]) {
      if (j !== goal && (cond[j] <= PURPLE_BOX || cond[j] === WALL)) target = j;
    }
    const tt = g.land[(target / N) | 0][target % N];
    const targetIsBlock = cond[target] <= PURPLE_BOX || cond[target] === WALL;
    const targetHits = targetIsBlock ? hitsToBreak(p.atk, tt) : 1;

    // ---- 足踏みの検出: 箱を壊す/叩く、敵を倒す、目的地にこれまでより近づく、のどれも無い手を数える
    let progressed = g.turn > prevTurn || living.length < prevAlive;
    if (targetIsBlock && target === prevTarget && tt.hp < prevTargetHp) progressed = true;
    const gd = fwd.dist[goal];
    if (!(gd >= (bestDist.get(goal) ?? Infinity))) { bestDist.set(goal, gd); progressed = true; }
    stall = progressed ? 0 : stall + 1;
    prevTurn = g.turn;
    prevAlive = living.length;
    prevTarget = target;
    prevTargetHp = tt.hp;

    // ---- 全体攻撃で経験値を稼ぐ（数回で全員倒せるときだけ）
    if (ah[0] > 0 && living.length >= 3) {
      const zd = Math.floor((p.atk * 0.9) / Math.max(1, g.monsters[0].def));
      const maxHp = living.reduce((mx, m) => Math.max(mx, m.hp), 0);
      const casts = zd > 0 ? Math.ceil(maxHp / zd) : Infinity;
      const incoming = local.reduce((sum, m) => sum + (cheb(m.x, m.y, p.x, p.y) <= 2 ? m.max : 0), 0);
      if ((casts === 1 || (casts <= 3 && ah[0] >= casts * 2)) && incoming * casts < p.hp * 0.5) return ['Z'];
    }

    // ---- 全消去を使う場面（紫箱は残るので、壁も敵も消えて一直線に進める）
    if (ah[2] > 0) {
      // 赤箱を掘らないと進めない
      const throughRed = cond[target] === RED_BOX && goal !== target;
      // もう壊す価値のある箱が無く、敵を避けるために大回りさせられている
      const plainStair = stair >= 0 ? dijkstra(plain, stop, start, false).dist[stair] : Infinity;
      const blocked = worth <= 2 && goal === stair && stairDist - plainStair > 10;
      if (throughRed || blocked || stall >= STALL_HARD) return ['C'];
    }
    if (stall >= STALL_HARD && living.length > 0) {
      if (ah[3] > 0) return ['D'];
      if (ah[4] > 0) return ['Enter'];
    }

    const field = dijkstra(cost, stop, target, true).dist;
    for (let i = 0; i < N * N; i++) if (field[i] === Infinity) field[i] = 5000;

    const hpBudget = p.hp;
    const deathCost = ah[5] > 0 ? 3000 : DEATH;
    // 進展が無いときは「待って倒す」をやめ、さらに続くなら多少の被弾も受け入れて進む
    const soft = stall >= STALL_SOFT;
    const killW = soft ? 0 : 2;
    const waitW = soft ? 2 : 0.3;
    const dmgScale = stall >= STALL_HARD ? 0.3 : 1;
    const best = tactical(g, { field, target, targetHits, local, cond, hpBudget, deathCost, killW, waitW, dmgScale });

    // ---- 避けきれない（死ぬ）ならコマンドで切り抜ける
    const doomed = !best || best.v >= deathCost - 100;
    const hurt = p.hp < p.maxHp * 0.35;
    const threatened = local.some((m) => cheb(m.x, m.y, p.x, p.y) <= 2);
    if (doomed || (hurt && threatened)) {
      if (ah[1] > 0 && p.hp < p.maxHp * 0.6 && !local.some((m) => m.max >= p.maxHp)) return ['X'];
      if (doomed) {
        // 箱がまだおいしい階では箱化（箱が増える）、そうでなければ全消去か次の階へ
        if (ah[3] > 0 && worth > 2) return ['D'];
        if (ah[2] > 0 && ah[2] >= ah[4]) return ['C'];
        if (ah[4] > 0) return ['Enter'];
        if (ah[2] > 0) return ['C'];
        if (ah[3] > 0) return ['D'];
        if (ah[0] > 0) return ['Z'];
      }
    }
    if (!best) return STEPS[(g.turn * 7) % 8][2];
    return STEPS[best.s][2];
  }

  // 祝福（アレンジルール）: 生存に効くものを優先し、ターンが苦しいときだけ時の祝福を取る
  // 添字は engine.js の BLESSINGS の並び
  function blessingKeys(g) {
    const b = g.blessing;
    const prio = [3, 8, 5, 1, 4, 6, 7, 2];
    if (g.turn < 150) prio[3] = 9;
    if (g.abilityHp[5] === 0) prio[5] = 6.5;
    let best = 0;
    for (let k = 1; k < 3; k++) if (prio[b.options[k]] > prio[b.options[best]]) best = k;
    return b.cursor === best ? ['Enter'] : ['Right'];
  }

  return {
    // 今フレームに押すキーの配列を返す
    nextKeys(g) {
      if (g.mode !== lastMode) {
        lastMode = g.mode;
        wait = g.mode === GAME_OVER || g.mode === GAME_CLEAR ? RESULT_WAIT : MENU_WAIT;
      }
      switch (g.mode) {
        case DUNGEON:
          return g.player.alive ? dungeonKeys(g) : [];
        case ENTRANCE:
        case OPTIONS:
          if (--wait > 0) return [];
          return ['Enter'];
        case BLESSING:
          return blessingKeys(g);
        case GAME_OVER:
        case GAME_CLEAR:
        case HOWTO_PLAY:
        case MUSEUM:
          if (--wait > 0) return [];
          return ['X'];
        default:
          return [];
      }
    },
    reset() {
      lastMode = -1;
    },
  };
}
