import { state } from './state.js';
import { GameMode } from './constants.js';
import { loadSprites, loadMap } from './mapLoader.js';
import { initAudio } from './audio.js';
import { initInput } from './input.js';
import { draw } from './renderer.js';
import { landSet, monsterSet, wordSet, movement, statusCheck, musicCheck, loadRecords } from './engine.js';
import { autoTick } from './autoPlay.js';

const canvas = document.getElementById('gameCanvas');
const ctx = canvas.getContext('2d');

const TICK_DELAY = 50;

function loop() {
  tick();
  setTimeout(loop, TICK_DELAY);
}

async function init() {
  initAudio();
  initInput();
  loadRecords();

  await Promise.all([loadSprites(), loadMap()]);

  monsterSet();
  landSet();
  wordSet();

  loop();
}

function tick() {
  const gm = state.gameMode;

  if (state.autoPlay) autoTick();

  if (gm === GameMode.Dungeon) {
    movement();
    statusCheck(false);
    wordSet();
  } else {
    movement();
    wordSet();
  }
  musicCheck();
  draw(ctx);
}

init();
