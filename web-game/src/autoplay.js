// オートプレイ（ブラウザ版の追加機能。VB6 版には無い）。
// 毎フレーム tick() の前に nextKeys(g) を呼び、返ってきたキーを keyDown() に渡す。
// DOM 非依存なので test/headless.js からも使える。

import {
  ENTRANCE, DUNGEON, HOWTO_PLAY, OPTIONS, GAME_OVER, GAME_CLEAR, MUSEUM,
  BLUE_BOX, RED_BOX, YELLOW_BOX, GREEN_BOX, PURPLE_BOX, STAIR, WALL, ROOM, ENEMY,
  LAND_NUMBER, WALL_BREAK, STEALTH, TURN_CONST,
} from './constants.js';

const N = LAND_NUMBER;
const MENU_WAIT = 30; // メニュー画面で次のキーを押すまでのフレーム数（1.5 秒）
const RESULT_WAIT = 80; // ゲームオーバー/クリア画面を見せるフレーム数（4 秒）

// 8 方向と、それを入力するキー
const STEPS = [
  [0, -1, ['Up']], [0, 1, ['Down']], [1, 0, ['Right']], [-1, 0, ['Left']],
  [1, -1, ['Up', 'Right']], [1, 1, ['Down', 'Right']],
  [-1, 1, ['Down', 'Left']], [-1, -1, ['Up', 'Left']],
];

const FIGHT_RISK = 1.0; // test/autoplay.js で 0.3〜3 を比べて一番深くまで潜れた値

const isEdge = (x, y) => x === 0 || y === 0 || x === N - 1 || y === N - 1;

// 1 回の攻撃のおおよそのダメージ（calcDamage の期待値）
const hitDamage = (atk, def) => Math.floor(atk / Math.max(1, def)) + 1;
const hitsToBreak = (atk, t) => Math.ceil(Math.max(1, t.hp) / hitDamage(atk, t.def));

// 箱を開けたときのおおよその価値（マス数換算）。ターン +10 の分は別に足す。
const BOX_VALUE = {
  [BLUE_BOX]: 6,
  [RED_BOX]: -40,
  [YELLOW_BOX]: 4,
  [GREEN_BOX]: 14,
  [PURPLE_BOX]: 30,
};

// 倒しきるまでに受けるダメージ（other は周りの敵からの分）が HP の半分未満なら戦う
function safeToFight(p, m, other) {
  const hits = Math.ceil(m.hp / hitDamage(p.atk, m.def));
  const taken = (hits - 1) * (hitDamage(m.atk, p.def) + other);
  return taken < p.hp * FIGHT_RISK;
}

export function createAutoPlayer() {
  let wait = 0;
  let lastMode = -1;

  // 最小ヒープ（[cost, index]）
  function dijkstra(g, cost) {
    const dist = new Float64Array(N * N).fill(Infinity);
    const prev = new Int32Array(N * N).fill(-1);
    const heap = [];
    const push = (d, i) => {
      heap.push([d, i]);
      let k = heap.length - 1;
      while (k > 0) {
        const pk = (k - 1) >> 1;
        if (heap[pk][0] <= heap[k][0]) break;
        [heap[pk], heap[k]] = [heap[k], heap[pk]];
        k = pk;
      }
    };
    const pop = () => {
      const top = heap[0];
      const last = heap.pop();
      if (heap.length) {
        heap[0] = last;
        let k = 0;
        for (;;) {
          const l = k * 2 + 1;
          const r = l + 1;
          let m = k;
          if (l < heap.length && heap[l][0] < heap[m][0]) m = l;
          if (r < heap.length && heap[r][0] < heap[m][0]) m = r;
          if (m === k) break;
          [heap[m], heap[k]] = [heap[k], heap[m]];
          k = m;
        }
      }
      return top;
    };

    const p = g.player;
    const start = p.x * N + p.y;
    dist[start] = 0;
    push(0, start);
    while (heap.length) {
      const [d, i] = pop();
      if (d > dist[i]) continue;
      const x = (i / N) | 0;
      const y = i % N;
      // 箱・敵・階段は「そこに入る」ところで止まる（通り抜けはしない）
      if (i !== start && cost[i] !== 1) continue;
      for (const [dx, dy] of STEPS) {
        const nx = x + dx;
        const ny = y + dy;
        if (nx < 0 || ny < 0 || nx >= N || ny >= N) continue;
        const j = nx * N + ny;
        const c = cost[j];
        if (c === Infinity) continue;
        if (d + c < dist[j]) {
          dist[j] = d + c;
          prev[j] = i;
          push(d + c, j);
        }
      }
    }
    return { dist, prev, start };
  }

  // 各マスに入るコスト。箱/敵は壊す（倒す）までの攻撃回数 + 1。
  function costMap(g, monsterAt) {
    const p = g.player;
    const cost = new Float64Array(N * N);
    for (let x = 0; x < N; x++) {
      for (let y = 0; y < N; y++) {
        const t = g.land[x][y];
        const i = x * N + y;
        const c = t.condition;
        if (c === ROOM || c === STAIR) cost[i] = 1;
        else if (c <= PURPLE_BOX) cost[i] = hitsToBreak(p.atk, t) + 1;
        else if (c === WALL) {
          cost[i] = p.condition === WALL_BREAK && !isEdge(x, y) ? hitsToBreak(p.atk, t) + 1 : Infinity;
        } else if (c === ENEMY) {
          const m = monsterAt.get(i);
          cost[i] = m ? Math.ceil(m.hp / hitDamage(p.atk, m.def)) + 1 : 1;
        } else cost[i] = 1;
      }
    }
    return cost;
  }

  // モンスターの周り（次の手で殴られうるマス）を通りにくくする
  function addDanger(g, cost, living) {
    const p = g.player;
    for (const m of living) {
      if (Math.max(Math.abs(m.x - p.x), Math.abs(m.y - p.y)) > 8) continue;
      if (safeToFight(p, m, 0)) continue; // 安全に倒せる相手は怖くない
      const danger = 2 + Math.ceil((hitDamage(m.atk, p.def) / Math.max(1, p.hp)) * 30);
      for (let dx = -1; dx <= 1; dx++) {
        for (let dy = -1; dy <= 1; dy++) {
          const x = m.x + dx;
          const y = m.y + dy;
          if (x < 0 || y < 0 || x >= N || y >= N) continue;
          const i = x * N + y;
          if (cost[i] !== Infinity) cost[i] += danger;
        }
      }
      // 敵のマス自体は倒しきれないなら実質通れない
      cost[m.x * N + m.y] += 50;
    }
  }

  function firstStep(path, target) {
    let j = target;
    while (path.prev[j] !== path.start && path.prev[j] !== -1) j = path.prev[j];
    if (path.prev[j] === -1) return null;
    return j;
  }

  function stepKeys(g, j) {
    const dx = ((j / N) | 0) - g.player.x;
    const dy = (j % N) - g.player.y;
    const s = STEPS.find(([sx, sy]) => sx === dx && sy === dy);
    return s ? s[2] : [];
  }

  function dungeonKeys(g) {
    const p = g.player;
    const ah = g.abilityHp;
    const living = [];
    const monsterAt = new Map();
    for (let i = 0; i < g.monsterNumber; i++) {
      const m = g.monsters[i];
      if (m.alive) {
        living.push(m);
        monsterAt.set(m.x * N + m.y, m);
      }
    }
    // 1 手で受けるおおよその最大ダメージ（隣接モンスター全員から）
    const stealth = p.condition === STEALTH;
    const near = stealth ? [] : living.filter((m) => Math.max(Math.abs(m.x - p.x), Math.abs(m.y - p.y)) <= 2);
    const incoming = near.reduce((s, m) => s + hitDamage(m.atk, p.def), 0);

    // --- 危ないときはコマンドを使う
    if (p.hp <= incoming * 2 || p.hp <= p.maxHp / 4) {
      if (ah[1] > 0 && p.hp < p.maxHp / 2) return ['X']; // Hp 全快
      if (near.length >= 2 && ah[3] > 0) return ['D']; // モンスター箱化
      if (near.length >= 1 && ah[0] > 0) return ['Z']; // 全体攻撃
    }

    const cost = costMap(g, monsterAt);
    if (!stealth) addDanger(g, cost, living);
    const path = dijkstra(g, cost);
    const sx = (() => {
      for (let x = 0; x < N; x++) for (let y = 0; y < N; y++) if (g.land[x][y].condition === STAIR) return x * N + y;
      return -1;
    })();
    const stairDist = sx >= 0 ? path.dist[sx] : Infinity;

    // ターンが足りなくなりそうなら「次の階へ」を使う
    if (g.turn <= 3 && ah[4] > 0) return ['Enter'];
    if (stairDist === Infinity && ah[4] > 0) return ['Enter'];

    // --- 隣のモンスターは 1 撃で倒せるときだけ倒す。
    // 倒しきれないと反撃を受けるので、それ以外は逃げる（敵はプレイヤーと同じ速さでしか追ってこない）
    if (!stealth) {
      let best = null;
      for (const [dx, dy] of STEPS) {
        const m = monsterAt.get((p.x + dx) * N + (p.y + dy));
        if (!m) continue;
        const hits = Math.ceil(m.hp / hitDamage(p.atk, m.def));
        if (safeToFight(p, m, incoming) && (!best || hits < best.hits)) best = { hits, j: m.x * N + m.y };
      }
      if (best) return stepKeys(g, best.j);
    }

    // --- 目的地選び: (かかる手数) − (価値) が最小のマス
    const turnGain = p.condition === TURN_CONST ? 0 : 10;
    const hungry = g.turn < stairDist + 40; // ターンが心もとない
    const hpRatio = p.hp / Math.max(1, p.maxHp);
    let target = -1;
    let bestScore = Infinity;
    if (sx >= 0 && stairDist < Infinity) {
      target = sx;
      bestScore = stairDist - (hungry ? 0 : 25);
    }
    for (let x = 0; x < N; x++) {
      for (let y = 0; y < N; y++) {
        const i = x * N + y;
        const c = g.land[x][y].condition;
        if (c > PURPLE_BOX || path.dist[i] === Infinity) continue;
        let value = BOX_VALUE[c] + turnGain * (hungry ? 1.5 : 0.3);
        if (c === BLUE_BOX && hpRatio < 0.5) value += 10;
        if (c === RED_BOX && !hungry) continue;
        const score = path.dist[i] - value;
        if (score < bestScore) { bestScore = score; target = i; }
      }
    }

    if (target < 0) {
      if (ah[2] > 0) return ['C']; // 全消去で道を開く
      return [STEPS[(g.turn * 7) % 8][2]].flat();
    }
    const j = firstStep(path, target);
    return j === null ? [] : stepKeys(g, j);
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
          if (--wait > 0) return [];
          return ['Enter'];
        case GAME_OVER:
        case GAME_CLEAR:
        case HOWTO_PLAY:
        case MUSEUM:
          if (--wait > 0) return [];
          return ['X'];
        case OPTIONS:
          if (--wait > 0) return [];
          return ['Enter'];
        default:
          return [];
      }
    },
    reset() {
      lastMode = -1;
    },
  };
}
