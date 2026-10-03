import { g, initGame, keyDown, tick } from './engine.js';
import { loadAssets } from './assets.js';
import { createRenderer } from './renderer.js';
import { createAudio } from './audio.js';
import { createAutoPlayer } from './autoplay.js';
import { createSfx } from './sfx.js';
import { DUNGEON } from './constants.js';

const SAVE_KEY = 'dungeon_and_box_save';
const RECORDS_KEY = 'dungeon_and_box_records';
const TICK_MS = 50; // VB6 の Timer1.Interval
// P キーで切り替えるオートプレイ: [ラベル, 1 フレームあたりの倍速]
const AUTO_MODES = [null, ['AUTO', 1], ['AUTO ×4', 4]];

// VB6 版の Data ファイル（ImscLib12.WriteText/ReadText）の代わり
function readJson(key) {
  try {
    const s = localStorage.getItem(key);
    return s ? JSON.parse(s) : null;
  } catch {
    return null;
  }
}

function writeJson(key, data) {
  try {
    localStorage.setItem(key, JSON.stringify(data));
  } catch {
    // 保存できない環境（プライベートモード等）では何もしない
  }
}

const storage = {
  load: () => readJson(SAVE_KEY),
  save: (data) => writeJson(SAVE_KEY, data),
  // 冒険の記録とルール設定（ブラウザ版独自）
  loadRecords: () => readJson(RECORDS_KEY),
  saveRecords: (data) => writeJson(RECORDS_KEY, data),
};

// KeyboardEvent.code → VB6 の KeyCode 名。code を使うので IME の状態に左右されない。
const KEY_MAP = {
  ArrowUp: 'Up', ArrowDown: 'Down', ArrowLeft: 'Left', ArrowRight: 'Right',
  Enter: 'Enter', NumpadEnter: 'Enter',
  KeyZ: 'Z', KeyX: 'X', KeyC: 'C', KeyD: 'D', KeyS: 'S', KeyA: 'A', KeyL: 'L',
};

async function main() {
  const canvas = document.getElementById('screen');
  const status = document.getElementById('status');
  const { sheets, map } = await loadAssets();
  const renderer = createRenderer(canvas, sheets);
  const audio = createAudio();
  const sfx = createSfx();
  const unlock = () => { audio.unlock(); sfx.unlock(); };

  initGame(map, storage);
  const auto = createAutoPlayer();
  let autoMode = 0;
  let muted = false;
  const showStatus = () => {
    const parts = [];
    if (AUTO_MODES[autoMode]) parts.push(`▶ ${AUTO_MODES[autoMode][0]}（Pで切替）`);
    if (muted) parts.push('♪ OFF');
    status.textContent = parts.join('　');
  };
  showStatus();

  window.addEventListener('keydown', (e) => {
    unlock();
    if (e.code === 'KeyM') {
      muted = audio.toggleMute();
      sfx.setMuted(muted);
      showStatus();
      return;
    }
    if (e.code === 'KeyP') {
      autoMode = (autoMode + 1) % AUTO_MODES.length;
      auto.reset();
      showStatus();
      return;
    }
    const key = KEY_MAP[e.code];
    if (!key) return;
    e.preventDefault();
    keyDown(key);
  });
  canvas.addEventListener('pointerdown', unlock);

  let last = performance.now();
  let acc = 0;
  function frame(now) {
    const dt = Math.min(100, now - last);
    acc += now - last;
    last = now;
    if (acc > TICK_MS * 10) acc = TICK_MS; // タブ復帰時に一気に進めない
    const speed = AUTO_MODES[autoMode]?.[1] ?? 1;
    while (acc >= TICK_MS) {
      for (let k = 0; k < speed; k++) {
        if (autoMode) {
          if (g.mode === DUNGEON) g.stats.auto = true; // オートプレイの結果は自己ベストに数えない
          for (const key of auto.nextKeys(g)) keyDown(key);
        }
        tick();
      }
      acc -= TICK_MS;
    }
    const events = g.events.splice(0);
    audio.sync(g);
    sfx.play(events);
    renderer.render(g, events, dt);
    requestAnimationFrame(frame);
  }
  requestAnimationFrame(frame);
}

main().catch((err) => {
  document.getElementById('status').textContent = `読み込みに失敗しました: ${err.message}`;
  console.error(err);
});
