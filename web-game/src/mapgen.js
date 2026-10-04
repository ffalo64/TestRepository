// フロアのランダム生成（ブラウザ版独自・アレンジルール用。VB6 版は Map40.bmp の 64 枚から選ぶ）。
// DOM 非依存。乱数は vb.js の rnd() を使うので、テストでは setRandom() で再現できる。
//
// 決まりごと
// - 紫箱と階段は 1 フロアに 1 個ずつ。どちらも、スタート地点から「床と赤以外の箱」だけを
//   通って行ける場所に置く（赤箱を壊さないと辿り着けない、ということは起きない）。
// - 序盤は青箱・緑箱が多くモンスターも遠い。深い階ほど赤箱の多いフロアが混ざる。
// - 形は VB6 版のマップの作風（広間、箱の輪、箱ブロック、箱の扉で仕切った小部屋、同心の回廊）を
//   型として持ち、寸法や配置を毎回変える。

import {
  BLUE_BOX, RED_BOX, YELLOW_BOX, GREEN_BOX, PURPLE_BOX, STAIR, WALL, ROOM, LAND_NUMBER,
} from './constants.js';
import { vbInt, rnd } from './vb.js';

const N = LAND_NUMBER;
const MIN_OPEN = 200; // 歩けるマスがこれより少ないフロアは作り直す

const ri = (n) => vbInt(rnd() * n); // 0..n-1
const rr = (a, b) => a + ri(b - a + 1); // a..b
const chance = (p) => rnd() < p;
const at = (x, y) => x * N + y;
const inner = (x, y) => x > 0 && y > 0 && x < N - 1 && y < N - 1;
const inset = (x, y) => Math.min(x, y, N - 1 - x, N - 1 - y);

function pickWeighted(weights) {
  let sum = 0;
  for (const w of weights) sum += w;
  let r = rnd() * sum;
  for (let k = 0; k < weights.length; k++) {
    r -= weights[k];
    if (r < 0) return k;
  }
  return weights.length - 1;
}

// ---------------------------------------------------------------- 箱の色のテーマ

// w: 箱の色の重み [青, 赤, 黄, 緑]、boxes: 箱の最低数 [min, max]、text: 階に着いたときの一言
const THEMES = {
  gentle: { w: [4, 0.4, 2, 4], boxes: [30, 45], text: '' },
  normal: { w: [3, 2, 2.5, 2.5], boxes: [20, 40], text: '' },
  spring: { w: [8, 0.4, 1, 1], boxes: [25, 45], text: '泉の階: 青箱が多い。' },
  bounty: { w: [1, 0.4, 1, 8], boxes: [25, 45], text: '恵みの階: 緑箱が多い。' },
  wonder: { w: [1, 1, 8, 1], boxes: [25, 45], text: '不思議の階: 黄箱が多い。' },
  harsh: { w: [2, 4.5, 2.5, 1.5], boxes: [20, 40], text: '' },
  cursed: { w: [0.4, 12, 0.8, 0.4], boxes: [70, 130], text: '災いの階: 赤箱だらけだ…' },
};

// 階層ごとのテーマの出やすさ。序盤は有利なものだけ、深い階ほど赤箱の階が増える
const THEME_TABLE = [
  [5, { gentle: 70, spring: 15, bounty: 15 }],
  [10, { gentle: 50, normal: 20, spring: 10, bounty: 10, wonder: 10 }],
  [30, { normal: 55, spring: 8, bounty: 8, wonder: 12, harsh: 12, cursed: 5 }],
  [100, { normal: 40, spring: 6, bounty: 6, wonder: 12, harsh: 21, cursed: 15 }],
  [Infinity, { normal: 35, spring: 5, bounty: 5, wonder: 10, harsh: 25, cursed: 20 }],
];

function pickTheme(floor) {
  const table = THEME_TABLE.find(([f]) => floor <= f)[1];
  const names = Object.keys(table);
  return THEMES[names[pickWeighted(names.map((n) => table[n]))]];
}

// スタート地点からこのマス数以内にはモンスターを置かない
function safeRadius(floor) {
  if (floor <= 10) return 8;
  if (floor <= 30) return 5;
  return 3;
}

// ---------------------------------------------------------------- 形づくり
// 各 layout は壁で埋まった c を掘り、箱を置く関数 (deco) => void を返す。
// deco: { put(x, y, color), color(), group(tiles), any() }

function carveRect(c, x0, y0, x1, y1) {
  for (let x = x0; x <= x1; x++) {
    for (let y = y0; y <= y1; y++) if (inner(x, y)) c[at(x, y)] = ROOM;
  }
}

function rectTiles(x0, y0, x1, y1) {
  const tiles = [];
  for (let x = x0; x <= x1; x++) for (let y = y0; y <= y1; y++) tiles.push([x, y]);
  return tiles;
}

function ringTiles(x0, y0, x1, y1) {
  return rectTiles(x0, y0, x1, y1).filter(([x, y]) => x === x0 || x === x1 || y === y0 || y === y1);
}

// 部屋と通路（ローグライクの定番）。部屋ごとに箱の置き方を変え、通路にもたまに箱が詰まっている
function layoutRooms(c) {
  const rooms = [];
  const want = rr(6, 10);
  for (let k = 0; k < 80 && rooms.length < want; k++) {
    const w = rr(4, 11);
    const h = rr(4, 9);
    const x = rr(1, N - 1 - w);
    const y = rr(1, N - 1 - h);
    if (rooms.some((r) => x <= r.x + r.w && r.x <= x + w && y <= r.y + r.h && r.y <= y + h)) continue;
    rooms.push({ x, y, w, h });
    carveRect(c, x, y, x + w - 1, y + h - 1);
  }
  const corridor = []; // 部屋の外を掘ったマス
  const dig = (x, y) => {
    if (c[at(x, y)] === WALL) { c[at(x, y)] = ROOM; corridor.push([x, y]); }
  };
  // 部屋の中の 1 点どうしを L 字の通路でつなぐ
  const link = (a, b) => {
    const ax = a.x + ri(a.w);
    const ay = a.y + ri(a.h);
    const bx = b.x + ri(b.w);
    const by = b.y + ri(b.h);
    const [cx, cy] = chance(0.5) ? [ax, by] : [bx, ay]; // L 字の角
    for (let x = Math.min(ax, bx); x <= Math.max(ax, bx); x++) dig(x, cy);
    for (let y = Math.min(ay, by); y <= Math.max(ay, by); y++) dig(cx, y);
  };
  rooms.sort((a, b) => a.x - b.x);
  for (let i = 0; i + 1 < rooms.length; i++) link(rooms[i], rooms[i + 1]);
  for (let k = rr(1, 3); k > 0 && rooms.length > 2; k--) link(rooms[ri(rooms.length)], rooms[ri(rooms.length)]);

  return (deco) => {
    for (const r of rooms) {
      const x1 = r.x + r.w - 1;
      const y1 = r.y + r.h - 1;
      const kind = rnd();
      if (kind < 0.2) continue;
      if (kind < 0.6) {
        for (let k = rr(1, 5); k > 0; k--) deco.put(rr(r.x, x1), rr(r.y, y1), deco.color());
      } else if (kind < 0.85) {
        deco.group(rectTiles(r.x + 1, r.y + 1, x1 - 1, y1 - 1)); // 周りを 1 マス空けた箱の山
      } else {
        deco.group(ringTiles(r.x + 1, r.y + 1, x1 - 1, y1 - 1)); // 箱の輪
      }
    }
    for (const [x, y] of corridor) if (c[at(x, y)] === ROOM && chance(0.04)) deco.put(x, y, deco.color());
  };
}

// 大広間。箱の輪・箱ブロックの格子・箱の列・柱と散らばった箱 のどれか
function layoutHall(c) {
  const x0 = rr(1, 6);
  const y0 = rr(1, 6);
  const x1 = N - 1 - rr(1, 6);
  const y1 = N - 1 - rr(1, 6);
  carveRect(c, x0, y0, x1, y1);
  const kind = ri(4);
  if (kind === 3) {
    // 柱
    const step = rr(4, 7);
    for (let x = x0 + 2; x < x1 - 1; x += step) {
      for (let y = y0 + 2; y < y1 - 1; y += step) {
        if (chance(0.8)) { c[at(x, y)] = WALL; if (chance(0.5)) c[at(x + 1, y)] = c[at(x, y + 1)] = c[at(x + 1, y + 1)] = WALL; }
      }
    }
  }
  return (deco) => {
    if (kind === 0) {
      // 同心の箱の輪（輪ごとに 1 色になりやすい）
      const step = rr(1, 3);
      let d = rr(2, 6);
      for (let k = rr(1, 4); k > 0; k--, d += step) {
        if (x1 - x0 - 2 * d < 3 || y1 - y0 - 2 * d < 3) break;
        deco.group(ringTiles(x0 + d, y0 + d, x1 - d, y1 - d));
      }
    } else if (kind === 1) {
      // 2x2 の箱ブロックの格子。外側からの距離ごとに色が揃う
      const colors = [];
      for (let x = x0 + 1; x + 1 <= x1 - 1; x += 4) {
        for (let y = y0 + 1; y + 1 <= y1 - 1; y += 4) {
          if (chance(0.15)) continue;
          const ring = Math.min(x - x0, y - y0, x1 - 1 - x, y1 - 1 - y) >> 2;
          colors[ring] ??= deco.color();
          for (const [bx, by] of rectTiles(x, y, x + 1, y + 1)) deco.put(bx, by, colors[ring]);
        }
      }
    } else if (kind === 2) {
      // 箱の列（ところどころ切れている）
      const step = rr(3, 6);
      const vertical = chance(0.5);
      const from = vertical ? x0 : y0;
      const to = vertical ? x1 : y1;
      for (let line = from + rr(1, 3); line < to; line += step) {
        const tiles = [];
        const a = (vertical ? y0 : x0) + rr(0, 3);
        const b = (vertical ? y1 : x1) - rr(0, 3);
        const gap = rr(a, b);
        for (let k = a; k <= b; k++) if (k !== gap || chance(0.4)) tiles.push(vertical ? [line, k] : [k, line]);
        deco.group(tiles);
      }
    } else {
      for (let k = rr(25, 60); k > 0; k--) deco.put(...deco.any(), deco.color());
    }
  };
}

// 洞窟（セルオートマトン）。箱は数個ずつ固まって落ちている
function layoutCave(c) {
  let a = new Uint8Array(N * N);
  for (let x = 0; x < N; x++) for (let y = 0; y < N; y++) a[at(x, y)] = !inner(x, y) || chance(0.42) ? 1 : 0;
  for (let it = 0; it < 5; it++) {
    const b = new Uint8Array(N * N);
    for (let x = 0; x < N; x++) {
      for (let y = 0; y < N; y++) {
        let walls = 0;
        for (let dx = -1; dx <= 1; dx++) {
          for (let dy = -1; dy <= 1; dy++) walls += inner(x + dx, y + dy) ? a[at(x + dx, y + dy)] : 1;
        }
        b[at(x, y)] = !inner(x, y) || walls >= 5 ? 1 : 0;
      }
    }
    a = b;
  }
  for (let i = 0; i < N * N; i++) c[i] = a[i] ? WALL : ROOM;
  return (deco) => {
    for (let k = rr(4, 9); k > 0; k--) {
      let [x, y] = deco.any();
      const tiles = [];
      for (let n = rr(3, 10); n > 0; n--) {
        tiles.push([x, y]);
        const nx = x + rr(-1, 1);
        const ny = y + rr(-1, 1);
        if (inner(nx, ny)) { x = nx; y = ny; }
      }
      deco.group(tiles);
    }
    for (let k = rr(5, 15); k > 0; k--) deco.put(...deco.any(), deco.color());
  };
}

// 壁で仕切った小部屋。仕切りには必ず扉があり、扉は箱で塞がっていることが多い
function layoutGrid(c) {
  const cuts = () => {
    const lines = [0];
    for (let p = rr(5, 10); p <= N - 6; p += rr(5, 10)) lines.push(p);
    lines.push(N - 1);
    return lines;
  };
  const xs = cuts();
  const ys = cuts();
  const doors = [];
  const cells = [];
  for (let i = 0; i + 1 < xs.length; i++) {
    for (let j = 0; j + 1 < ys.length; j++) {
      const cell = [xs[i] + 1, ys[j] + 1, xs[i + 1] - 1, ys[j + 1] - 1];
      carveRect(c, ...cell);
      cells.push(cell);
      if (i > 0) doors.push([xs[i], rr(cell[1], cell[3])]);
      if (j > 0) doors.push([rr(cell[0], cell[2]), ys[j]]);
    }
  }
  for (const [x, y] of doors) c[at(x, y)] = ROOM;
  return (deco) => {
    for (const [x, y] of doors) if (chance(0.6)) deco.put(x, y, deco.color());
    for (const [x0, y0, x1, y1] of cells) {
      const kind = rnd();
      if (kind < 0.3) continue;
      if (kind < 0.65) {
        for (let k = rr(1, 4); k > 0; k--) deco.put(rr(x0, x1), rr(y0, y1), deco.color());
      } else if (kind < 0.85) {
        deco.group(rectTiles(x0, y0, x1, y1)); // 箱で埋まった部屋
      } else {
        deco.group(rectTiles(x0, y0, x1, y1).filter(([x, y]) => (x + y) % 2 === 0)); // 市松
      }
    }
  };
}

// 同心の回廊。回廊どうしは箱の門でつながり、中心に小部屋がある
function layoutRings(c) {
  const s = rr(2, 4); // 回廊の幅 s-1 と壁 1
  const gates = [];
  for (let d = 1; ; d += s) {
    if (N - 2 * d < s + 5) {
      carveRect(c, d, d, N - 1 - d, N - 1 - d);
      break;
    }
    const w = d + s - 1; // この回廊の内側の壁
    for (let x = d; x <= N - 1 - d; x++) {
      for (let y = d; y <= N - 1 - d; y++) if (inset(x, y) < w) c[at(x, y)] = ROOM;
    }
    for (let k = rr(1, 3); k > 0; k--) {
      const pos = rr(w + 1, N - 2 - w);
      const gate = [[pos, w], [pos, N - 1 - w], [w, pos], [N - 1 - w, pos]][ri(4)];
      c[at(gate[0], gate[1])] = ROOM;
      gates.push(gate);
    }
  }
  return (deco) => {
    for (const [x, y] of gates) if (chance(0.7)) deco.put(x, y, deco.color());
    for (let k = rr(8, 25); k > 0; k--) deco.put(...deco.any(), deco.color());
    const m = N >> 1;
    if (chance(0.6)) deco.group(rectTiles(m - 1, m - 1, m, m));
  };
}

const LAYOUTS = [layoutRooms, layoutHall, layoutCave, layoutGrid, layoutRings];
const LAYOUT_WEIGHTS = [35, 20, 20, 15, 10];

// ---------------------------------------------------------------- 仕上げ

// 一番大きい（上下左右につながった）床のかたまりだけを残し、そのマスの一覧を返す
function keepLargest(c) {
  const label = new Int32Array(N * N).fill(-1);
  let best = [];
  for (let s = 0; s < N * N; s++) {
    if (c[s] !== ROOM || label[s] >= 0) continue;
    const comp = [s];
    label[s] = s;
    for (let k = 0; k < comp.length; k++) {
      const i = comp[k];
      for (const j of [i - N, i + N, i - 1, i + 1]) {
        if (c[j] === ROOM && label[j] < 0) { label[j] = s; comp.push(j); }
      }
    }
    if (comp.length > best.length) best = comp;
  }
  const keep = best.length ? label[best[0]] : -2;
  for (let i = 0; i < N * N; i++) if (c[i] === ROOM && label[i] !== keep) c[i] = WALL;
  return best;
}

// start から各 target まで「赤箱を壊さずに」行けるようにする。
// 赤箱をなるべく避けた経路を求め、それでも経路上に残った赤箱は別の色に塗り替える。
function ensureReachable(c, start, targets, nonRed) {
  const enter = (v) => (v === ROOM || v === STAIR ? 1 : v === RED_BOX ? 1000 : v <= PURPLE_BOX ? 6 : Infinity);
  const dist = new Float64Array(N * N).fill(Infinity);
  const prev = new Int32Array(N * N).fill(-1);
  const done = new Uint8Array(N * N);
  dist[start] = 0;
  for (;;) {
    let i = -1;
    for (let k = 0; k < N * N; k++) if (!done[k] && dist[k] < Infinity && (i < 0 || dist[k] < dist[i])) i = k;
    if (i < 0) break;
    done[i] = 1;
    if (targets.includes(i)) continue; // 階段・紫箱の先へは進まない
    const x = (i / N) | 0;
    const y = i % N;
    for (let dx = -1; dx <= 1; dx++) {
      for (let dy = -1; dy <= 1; dy++) {
        const j = at(x + dx, y + dy); // 外周は必ず壁なので範囲外にはならない
        const d = dist[i] + enter(c[j]);
        if (d < dist[j]) { dist[j] = d; prev[j] = i; }
      }
    }
  }
  for (const t of targets) {
    for (let j = prev[t]; j >= 0 && j !== start; j = prev[j]) if (c[j] === RED_BOX) c[j] = nonRed();
  }
}

// フロアを 1 つ作る。
// 戻り値: { cond: マスごとの Condition (x * LAND_NUMBER + y), x, y: スタート地点,
//           safe: モンスターを置かない半径, text: 階に着いたときの一言 }
export function generateFloor(floor) {
  const theme = pickTheme(floor);
  let c;
  let open;
  let decorate;
  for (let attempt = 0; ; attempt++) {
    c = new Int8Array(N * N).fill(WALL);
    // 何度作っても狭いときは、必ず広くなる大広間にする
    const layout = attempt < 8 ? LAYOUTS[pickWeighted(LAYOUT_WEIGHTS)] : layoutHall;
    decorate = layout(c);
    open = keepLargest(c);
    if (open.length >= MIN_OPEN) break;
  }

  // スタート地点・階段・紫箱（同じかたまりの床から選ぶ）
  const pool = open.slice();
  const take = () => {
    const k = ri(pool.length);
    const i = pool[k];
    pool[k] = pool[pool.length - 1];
    pool.pop();
    return i;
  };
  const start = take();
  const stair = take();
  const purple = take();
  c[stair] = STAIR;
  c[purple] = PURPLE_BOX;

  // 箱
  let boxes = 0;
  const color = () => pickWeighted(theme.w);
  const put = (x, y, col) => {
    const i = at(x, y);
    if (!inner(x, y) || c[i] !== ROOM || i === start) return;
    c[i] = col;
    boxes++;
  };
  const any = () => {
    const i = open[ri(open.length)];
    return [(i / N) | 0, i % N];
  };
  const group = (tiles) => {
    const mono = chance(0.6);
    const col = color();
    for (const [x, y] of tiles) put(x, y, mono ? col : color());
  };
  decorate({ put, color, group, any });
  // 箱が少なすぎる階には散らして足す（箱はターンの元手なので）
  const want = Math.min(rr(theme.boxes[0], theme.boxes[1]), open.length >> 2);
  for (let k = 0; boxes < want && k < 1000; k++) put(...any(), color());

  const nonRedWeights = [theme.w[BLUE_BOX], 0, theme.w[YELLOW_BOX], theme.w[GREEN_BOX]];
  ensureReachable(c, start, [stair, purple], () => pickWeighted(nonRedWeights));

  return { cond: c, x: (start / N) | 0, y: start % N, safe: safeRadius(floor), text: theme.text };
}
