// 描画・演出・効果音の smoke テスト。Canvas と Web Audio をモックに差し替え、
// オートプレイで進めながら毎フレーム描画して、例外が出ないことだけを確かめる。
// 使い方: node test/render.js [ticks=6000]
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { decodeBmp } from '../src/bmp.js';
import { setRandom, mulberry32 } from '../src/vb.js';
import { initGame, keyDown, tick, g } from '../src/engine.js';
import { createAutoPlayer } from '../src/autoplay.js';
import { createRenderer } from '../src/renderer.js';
import { createSfx } from '../src/sfx.js';
import { BLESSING, GAME_OVER } from '../src/constants.js';

const TICKS = Number(process.argv[2] ?? 6000);

// 何を呼んでも何を読んでも通るモック（measureText などは数値を返す）
function mock(overrides = {}) {
  const fn = () => proxy;
  const proxy = new Proxy(fn, {
    get: (_, key) => (key in overrides ? overrides[key] : key === Symbol.toPrimitive ? () => 0 : proxy),
    set: () => true,
  });
  return proxy;
}

const calls = {};
const ctx = new Proxy({}, {
  get: (_, key) => {
    if (key === 'measureText') return () => ({ width: 100 });
    if (key === 'createRadialGradient') return () => ({ addColorStop() {} });
    return (...args) => {
      calls[key] = (calls[key] ?? 0) + 1;
      for (const a of args) if (typeof a === 'number' && Number.isNaN(a)) throw new Error(`${String(key)} に NaN`);
    };
  },
  set: () => true,
});
const canvas = { getContext: () => ctx, getBoundingClientRect: () => ({ width: 600 }), width: 0, height: 0 };

globalThis.window = {
  devicePixelRatio: 1,
  AudioContext: class {
    constructor() { this.currentTime = 0; this.sampleRate = 8000; this.state = 'running'; this.destination = {}; }
    createGain() { return mock(); }
    createOscillator() { return mock(); }
    createBiquadFilter() { return mock(); }
    createBufferSource() { return mock(); }
    createBuffer(_, len) { return { getChannelData: () => new Float32Array(len) }; }
    resume() {}
  },
};

const here = path.dirname(fileURLToPath(import.meta.url));
const buf = fs.readFileSync(path.join(here, '../public/images/Map40.bmp'));
const mapImg = decodeBmp(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));

setRandom(mulberry32(3));
let records = null;
initGame(mapImg, { load: () => null, save() {}, loadRecords: () => records, saveRecords: (r) => { records = r; } });
const renderer = createRenderer(canvas, { player: {}, monster: {}, box: {}, room: {}, wall: {} });
const sfx = createSfx();
sfx.unlock();
const ai = createAutoPlayer();

const seen = {};
const modes = new Set();
let sfxTime = 0;
for (let t = 0; t < TICKS; t++) {
  for (const k of ai.nextKeys(g)) keyDown(k);
  tick();
  const events = g.events.splice(0);
  for (const e of events) seen[e.t] = (seen[e.t] ?? 0) + 1;
  modes.add(g.mode);
  sfxTime += 0.05;
  sfx.play(events);
  renderer.render(g, events, 50);
}

const need = ['hit', 'kill', 'box', 'hurt', 'level', 'floor', 'bless', 'blessed', 'gameover'];
const missing = need.filter((k) => !seen[k]);
console.log('events', JSON.stringify(seen));
console.log('modes', [...modes].join(','), 'records', JSON.stringify(records?.runs?.[0] ?? null));
if (!modes.has(BLESSING) || !modes.has(GAME_OVER)) missing.push('mode');
if (!records?.runs?.length) missing.push('records');
if (!calls.fillText || !calls.drawImage) missing.push('draw calls');
console.log(missing.length ? `FAILED: 出なかったもの ${missing.join(', ')}` : 'OK');
process.exit(missing.length ? 1 : 0);
