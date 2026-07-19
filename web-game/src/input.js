import { GameMode } from './constants.js';
import { state } from './state.js';
import { landSet, wordSet, monsterSet, floorSet, abilityEffect, movement } from './engine.js';
import { onUserInteraction } from './audio.js';

export function initInput() {
  window.addEventListener('keydown', e => {
    onUserInteraction();
    handleKey(e.code);
    e.preventDefault();
  });
}

function handleKey(code) {
  const { player } = state;
  const gm = state.gameMode;

  if (gm === GameMode.Entrance) {
    if (code === 'KeyZ') {
      state.gameMode = GameMode.HowtoPlay;
      landSet(); monsterSet();
    } else if (code === 'KeyC') {
      state.gameMode = GameMode.Options;
    } else if (code === 'KeyD') {
      state.gameMode = GameMode.Museum;
      landSet();
    } else if (code === 'Enter') {
      state.gameMode = GameMode.Dungeon;
      landSet(); monsterSet(); floorSet();
    }

  } else if (gm === GameMode.Dungeon) {
    if (code === 'KeyP') {
      state.autoPlay = !state.autoPlay;
    } else if (code === 'ArrowUp')    player.direction *= 2;
    else if (code === 'ArrowDown')  player.direction *= 3;
    else if (code === 'ArrowRight') player.direction *= 5;
    else if (code === 'ArrowLeft')  player.direction *= 7;
    else if (['KeyZ','KeyX','KeyC','KeyD','KeyS','Enter','KeyA'].includes(code)) {
      abilityEffect(code === 'Enter' ? 'Enter' : code);
    }

  } else if (gm === GameMode.HowtoPlay) {
    if (code === 'ArrowRight') player.direction *= 5;
    else if (code === 'ArrowLeft') player.direction *= 7;
    else if (code === 'KeyX') {
      state.gameMode = GameMode.Entrance;
      landSet(); wordSet();
    }

  } else if (gm === GameMode.Options) {
    if (code === 'ArrowRight') player.direction *= 5;
    else if (code === 'ArrowLeft') player.direction *= 7;
    else if (code === 'KeyX') {
      state.gameMode = GameMode.Entrance;
      landSet(); wordSet();
    } else if (code === 'Enter') {
      state.gameMode = GameMode.Dungeon;
      landSet(); monsterSet(); floorSet();
    }

  } else if (gm === GameMode.GameOver || gm === GameMode.GameClear) {
    if (code === 'KeyX') {
      state.gameMode = GameMode.Entrance;
      landSet(); wordSet();
    }

  } else if (gm === GameMode.Museum) {
    if (code === 'ArrowRight') player.direction *= 5;
    else if (code === 'ArrowLeft') player.direction *= 7;
    else if (code === 'KeyX') {
      state.gameMode = GameMode.Entrance;
      landSet(); wordSet();
    } else if (code === 'KeyL') {
      abilityEffect('KeyL');
    }
  }
}
