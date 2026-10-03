// オートプレイのヘッドレス評価。使い方: node test/autoplay.js [games=20] [maxTicks=200000]
// 環境変数: GOAL=到達率を数える階 (既定 100), STOP=この階に着いたら打ち切る, SEED0=最初の seed,
//           ARRANGE=0 でオリジナルルール（祝福なし）
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { decodeBmp } from '../src/bmp.js';
import { setRandom, mulberry32 } from '../src/vb.js';
import { initGame, keyDown, tick, g } from '../src/engine.js';
import { createAutoPlayer } from '../src/autoplay.js';
import { DUNGEON, GAME_OVER, GAME_CLEAR } from '../src/constants.js';

const GAMES = Number(process.argv[2] ?? 20);
const MAX_TICKS = Number(process.argv[3] ?? 200000);
const GOAL = Number(process.env.GOAL ?? 100);
const STOP = Number(process.env.STOP ?? Infinity);
const SEED0 = Number(process.env.SEED0 ?? 1);
const here = path.dirname(fileURLToPath(import.meta.url));
const buf = fs.readFileSync(path.join(here, '../public/images/Map40.bmp'));
const mapImg = decodeBmp(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));

const floors = [];
let clears = 0;
const causes = {};
let failures = 0;
let reached = 0;
let stalls = 0;
let ticks = 0;
const t0 = Date.now();
for (let seed = SEED0; seed < SEED0 + GAMES; seed++) {
  setRandom(mulberry32(seed));
  initGame(mapImg, null);
  g.arrange = process.env.ARRANGE !== '0';
  const ai = createAutoPlayer();
  let maxFloor = 0;
  let result = 'timeout';
  let t = 0;
  // 足踏みの検出: 直近 20 手で 3 マス以内しか動かず、箱も壊さず敵も倒していない手を数える
  const recent = [];
  let stall = 0;
  let quiet = 0;
  const alive = () => { let n = 0; for (let i = 0; i < g.monsterNumber; i++) if (g.monsters[i].alive) n++; return n; };
  try {
    for (t = 1; t <= MAX_TICKS; t++) {
      const before = g.mode === DUNGEON ? [g.floor, g.turn, alive()] : null;
      for (const k of ai.nextKeys(g)) keyDown(k);
      tick();
      if (before && g.mode === DUNGEON) {
        const progressed = g.floor !== before[0] || g.turn > before[1] || alive() < before[2];
        quiet = progressed ? 0 : quiet + 1;
        recent.push(g.player.x * 100 + g.player.y);
        if (recent.length > 20) recent.shift();
        if (quiet >= 20 && new Set(recent).size <= 3) stall++;
      }
      if (g.mode === DUNGEON) maxFloor = Math.max(maxFloor, g.floor);
      if (maxFloor >= STOP) { result = 'stop'; break; }
      if (g.mode === GAME_OVER || g.mode === GAME_CLEAR) {
        result = g.mode === GAME_OVER ? (g.turn <= 0 ? 'turn0' : 'hp0') : 'clear';
        break;
      }
    }
  } catch (e) {
    failures++;
    result = 'error: ' + e.message;
  }
  if (result === 'clear') clears++;
  if (maxFloor >= GOAL || result === 'clear') reached++;
  causes[result] = (causes[result] ?? 0) + 1;
  floors.push(maxFloor);
  stalls += stall;
  ticks += t;
  console.log(`seed ${String(seed).padStart(3)}: ${result.padEnd(8)} floor=${String(maxFloor).padStart(4)} lv=${g.player.level} ticks=${t} stall=${stall} ah=${g.abilityHp.slice(0, 6).join(',')}`);
}
floors.sort((a, b) => a - b);
const avg = floors.reduce((s, f) => s + f, 0) / floors.length;
console.log(`---- avg floor ${avg.toFixed(1)}, median ${floors[floors.length >> 1]}, max ${floors.at(-1)}, clears ${clears}/${GAMES}, ${GOAL}F+ ${reached}/${GAMES} (${((reached / GAMES) * 100).toFixed(0)}%) stall ${((stalls / ticks) * 100).toFixed(1)}% ${JSON.stringify(causes)} ${((Date.now() - t0) / 1000).toFixed(1)}s`);
process.exit(failures ? 1 : 0);
