// ヘッドレス smoke / invariant テスト。使い方: node test/headless.js [seeds=20] [maxTicks=20000]
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { decodeBmp } from '../src/bmp.js';
import { cLng, setRandom, mulberry32 } from '../src/vb.js';
import { initGame, keyDown, tick, g } from '../src/engine.js';
import { generateFloor } from '../src/mapgen.js';
import {
  ENTRANCE, DUNGEON, GAME_OVER, GAME_CLEAR, MUSEUM, BLESSING, STAIR, ENEMY, LAND_NUMBER,
  BLUE_BOX, RED_BOX, YELLOW_BOX, GREEN_BOX, PURPLE_BOX, WALL, ROOM,
} from '../src/constants.js';

const SEEDS = Number(process.argv[2] ?? 20);
const MAX_TICKS = Number(process.argv[3] ?? 20000);
const here = path.dirname(fileURLToPath(import.meta.url));
const buf = fs.readFileSync(path.join(here, '../public/images/Map40.bmp'));
const mapImg = decodeBmp(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));

let failures = 0;
let blessings = 0; // 祝福の選択画面に居たフレーム数
function fail(msg) { failures++; console.error('FAIL: ' + msg); }

// ---- cLng 単体チェック
const z = (v) => (Object.is(v, -0) ? 0 : v);
const clngCases = [[2.5, 2], [3.5, 4], [17.500000000000004, 18], [-0.5, 0], [1.4, 1], [-1.5, -2], [0.5, 0]];
for (const [inp, exp] of clngCases) {
  const got = z(cLng(inp));
  if (got !== exp) fail(`cLng(${inp}) = ${got}, expected ${exp}`);
}

// ---- ランダム生成フロアの決まりごと（紫箱と階段が 1 個ずつあり、赤箱を壊さずに辿り着ける）
function checkFloor(gen, ctx) {
  const N = LAND_NUMBER;
  const c = gen.cond;
  const count = (v) => c.reduce((n, x) => n + (x === v ? 1 : 0), 0);
  if (count(STAIR) !== 1) fail(`${ctx}: 階段が ${count(STAIR)} 個`);
  if (count(PURPLE_BOX) !== 1) fail(`${ctx}: 紫箱が ${count(PURPLE_BOX)} 個`);
  for (let k = 0; k < N; k++) {
    if (c[k] !== WALL || c[(N - 1) * N + k] !== WALL || c[k * N] !== WALL || c[k * N + N - 1] !== WALL) {
      fail(`${ctx}: 外周が壁でない`);
      break;
    }
  }
  const start = gen.x * N + gen.y;
  if (c[start] !== ROOM) fail(`${ctx}: スタート地点が床でない (${c[start]})`);
  // 床と赤以外の箱だけを通る 8 方向の探索。階段・紫箱には入れるが、その先へは進まない
  const seen = new Uint8Array(N * N);
  const queue = [start];
  seen[start] = 1;
  for (let k = 0; k < queue.length; k++) {
    const i = queue[k];
    if (c[i] === STAIR || c[i] === PURPLE_BOX) continue;
    for (const d of [-N - 1, -N, -N + 1, -1, 1, N - 1, N, N + 1]) {
      const j = i + d;
      if (!seen[j] && c[j] !== WALL && c[j] !== RED_BOX) { seen[j] = 1; queue.push(j); }
    }
  }
  if (!seen[c.indexOf(STAIR)]) fail(`${ctx}: 階段に辿り着けない`);
  if (!seen[c.indexOf(PURPLE_BOX)]) fail(`${ctx}: 紫箱に辿り着けない`);
  return { red: count(RED_BOX), boxes: count(BLUE_BOX) + count(RED_BOX) + count(YELLOW_BOX) + count(GREEN_BOX) };
}

{
  setRandom(mulberry32(12345));
  const bands = [[1, 10], [11, 30], [31, 100], [101, 999]];
  const ratios = bands.map(([from, to]) => {
    let red = 0;
    let boxes = 0;
    for (let k = 0; k < 400; k++) {
      const floor = from + (k % (to - from + 1));
      const r = checkFloor(generateFloor(floor), `generateFloor(${floor}) #${k}`);
      red += r.red;
      boxes += r.boxes;
    }
    return red / boxes;
  });
  console.log('赤箱の割合 ' + bands.map(([from, to], k) => `${from}-${to}F: ${(ratios[k] * 100).toFixed(0)}%`).join(', '));
  if (!(ratios[0] < 0.1)) fail(`序盤の赤箱が多すぎる (${ratios[0]})`);
  if (!(ratios[0] < ratios[1] && ratios[1] < ratios[2])) fail(`赤箱の割合が階層とともに増えていない (${ratios})`);
}

const isInt = (v) => Number.isInteger(v);
const inRange = (v) => isInt(v) && v >= 0 && v < LAND_NUMBER;

function checkInvariants(ctx, stats) {
  const p = g.player;
  if (!inRange(p.x) || !inRange(p.y)) throw new Error(`${ctx}: player pos (${p.x},${p.y})`);
  for (const k of ['hp', 'maxHp', 'atk', 'def', 'exp', 'level']) {
    if (Number.isNaN(p[k]) || typeof p[k] !== 'number') throw new Error(`${ctx}: player.${k}=${p[k]}`);
  }
  if (Number.isNaN(g.turn) || Number.isNaN(g.floor)) throw new Error(`${ctx}: turn/floor NaN`);
  g.abilityHp.forEach((v, i) => {
    if (Number.isNaN(v) || typeof v !== 'number') throw new Error(`${ctx}: abilityHp[${i}]=${v}`);
  });
  for (let x = 0; x < LAND_NUMBER; x++) {
    for (let y = 0; y < LAND_NUMBER; y++) {
      const c = g.land[x][y].condition;
      if (!isInt(c) || c < 0 || c > 8) throw new Error(`${ctx}: land[${x}][${y}].condition=${c}`);
    }
  }
  const seen = new Map();
  for (let i = 0; i < g.monsterNumber; i++) {
    const m = g.monsters[i];
    if (!m.alive) continue;
    if (!inRange(m.x) || !inRange(m.y)) throw new Error(`${ctx}: monster ${i} pos (${m.x},${m.y})`);
    const key = m.x * 100 + m.y;
    const overlapsMon = seen.has(key);
    // 後続モンスターとの重なりも拾うため、先に全員の位置を登録してから判定する
    seen.set(key, (seen.get(key) ?? 0) + 1);
    void overlapsMon;
  }
  for (let i = 0; i < g.monsterNumber; i++) {
    const m = g.monsters[i];
    if (!m.alive) continue;
    const overlap = (m.x === p.x && m.y === p.y) || seen.get(m.x * 100 + m.y) > 1;
    if (g.land[m.x][m.y].condition !== ENEMY) {
      if (overlap) stats.overlaps++;
      else throw new Error(`${ctx}: monster ${i} at (${m.x},${m.y}) on tile cond ${g.land[m.x][m.y].condition}, floor ${g.floor}`);
    } else if (overlap) stats.overlaps++;
  }
}

function findStair() {
  for (let x = 0; x < LAND_NUMBER; x++) {
    for (let y = 0; y < LAND_NUMBER; y++) if (g.land[x][y].condition === STAIR) return [x, y];
  }
  return null;
}

function botKey(rand) {
  const r = rand();
  if (r < 0.003) return 'S';
  if (r < 0.023) return ['Z', 'X', 'C', 'D', 'Enter'][Math.floor(rand() * 5)];
  if (rand() < 0.3) return ['Up', 'Down', 'Left', 'Right'][Math.floor(rand() * 4)];
  const s = findStair();
  if (!s) return ['Up', 'Down', 'Left', 'Right'][Math.floor(rand() * 4)];
  const dx = s[0] - g.player.x;
  const dy = s[1] - g.player.y;
  if (Math.abs(dx) >= Math.abs(dy) && dx !== 0) return dx > 0 ? 'Right' : 'Left';
  if (dy !== 0) return dy > 0 ? 'Down' : 'Up';
  return dx > 0 ? 'Right' : 'Left';
}

const totals = { gameover: 0, clear: 0, timeout: 0, overlaps: 0, saves: 0, loads: 0 };
console.log(`seeds=${SEEDS} maxTicks=${MAX_TICKS}`);

for (let seed = 1; seed <= SEEDS; seed++) {
  const rand = mulberry32(seed * 7919);
  setRandom(mulberry32(seed));
  let saved = null;
  const god = seed % 2 === 0; // 偶数 seed は不死身(HP/ターン補充)にして深い階層まで検証
  const store = { load: () => saved, save: (o) => { saved = JSON.parse(JSON.stringify(o)); } };
  const stats = { overlaps: 0 };
  let result = 'timeout';
  let maxFloor = 0;
  let t = 0;
  try {
    initGame(mapImg, store);
    keyDown('Enter');
    if (g.mode !== DUNGEON) throw new Error(`seed ${seed}: Enter on title -> mode ${g.mode}`);
    for (t = 1; t <= MAX_TICKS; t++) {
      const ctx = `seed ${seed} tick ${t}`;
      if (god) { g.player.hp = g.player.maxHp; if (g.turn < 100) g.turn = 100; }
      keyDown(botKey(rand));
      tick();
      if (g.mode === DUNGEON) {
        checkInvariants(ctx, stats);
        maxFloor = Math.max(maxFloor, g.floor);
      } else if (g.mode === GAME_OVER || g.mode === GAME_CLEAR) {
        result = g.mode === GAME_OVER ? 'gameover' : 'clear';
        maxFloor = Math.max(maxFloor, g.floor);
        keyDown('X');
        break;
      } else if (g.mode === ENTRANCE) {
        if (!saved) throw new Error(`${ctx}: returned to ENTRANCE without save`);
        totals.saves++; if (process.env.V) console.log(`  save at tick ${t} seed ${seed} floor ${saved.floor} hp ${saved.player.hp}`);
        const savedFloor = saved.floor;
        keyDown('D');
        if (g.mode !== MUSEUM) throw new Error(`${ctx}: D on title -> mode ${g.mode}`);
        keyDown('L');
        // 10 の倍数の階でセーブすると、ロード直後に祝福の選択画面になる
        if (g.mode !== DUNGEON && g.mode !== BLESSING) throw new Error(`${ctx}: load -> mode ${g.mode}`);
        if (g.floor !== savedFloor + 1) throw new Error(`${ctx}: loaded floor ${g.floor}, saved ${savedFloor}`);
        totals.loads++;
        checkInvariants(`${ctx} (after load)`, stats);
      } else if (g.mode === BLESSING) {
        blessings++;
        keyDown('Enter'); // 祝福（アレンジルール）は真ん中をそのまま選ぶ
      } else {
        throw new Error(`${ctx}: unexpected mode ${g.mode}`);
      }
    }
  } catch (e) {
    fail(`seed ${seed} tick ${t}: ${e.message}`);
    result = 'error';
  }
  if (result in totals) totals[result]++;
  totals.overlaps += stats.overlaps;
  console.log(`seed ${String(seed).padStart(3)}: ${result.padEnd(8)} god=${seed % 2 === 0 ? 1 : 0} maxFloor=${String(maxFloor).padStart(4)} ticks=${t > MAX_TICKS ? MAX_TICKS : t} overlaps=${stats.overlaps}`);
}

console.log('---- summary');
console.log(JSON.stringify(totals));
if (!blessings) fail('祝福の選択画面が一度も出なかった');
console.log(failures ? `FAILED (${failures})` : 'OK');
process.exit(failures ? 1 : 0);
