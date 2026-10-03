// オートプレイのヘッドレス評価。使い方: node test/autoplay.js [games=20] [maxTicks=200000]
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
const here = path.dirname(fileURLToPath(import.meta.url));
const buf = fs.readFileSync(path.join(here, '../public/images/Map40.bmp'));
const mapImg = decodeBmp(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));

const floors = [];
let clears = 0;
const causes = {};
let failures = 0;
for (let seed = 1; seed <= GAMES; seed++) {
  setRandom(mulberry32(seed));
  initGame(mapImg, null);
  const ai = createAutoPlayer();
  let maxFloor = 0;
  let result = 'timeout';
  let t = 0;
  try {
    for (t = 1; t <= MAX_TICKS; t++) {
      for (const k of ai.nextKeys(g)) keyDown(k);
      tick();
      if (g.mode === DUNGEON) maxFloor = Math.max(maxFloor, g.floor);
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
  causes[result] = (causes[result] ?? 0) + 1;
  floors.push(maxFloor);
  console.log(`seed ${String(seed).padStart(3)}: ${result.padEnd(8)} floor=${String(maxFloor).padStart(4)} lv=${g.player.level} ticks=${t}`);
}
floors.sort((a, b) => a - b);
const avg = floors.reduce((s, f) => s + f, 0) / floors.length;
console.log(`---- avg floor ${avg.toFixed(1)}, median ${floors[floors.length >> 1]}, max ${floors.at(-1)}, clears ${clears}/${GAMES} ${JSON.stringify(causes)}`);
process.exit(failures ? 1 : 0);
