import { SC, LAND_NUMBER, GameMode } from './constants.js';

export function makeStatus() {
  return {
    left:0, top:0, oleft:0, otop:0,
    width:SC, height:SC,
    alive:false,
    hp:0, maxHp:0,
    exp:0, atk:0, def:1,
    level:1, condition:0,
    direction:1,
    name:'', explanation:'',
    ability:0
  };
}

function makeLandGrid() {
  const g = [];
  for (let i = 0; i < LAND_NUMBER; i++) {
    g[i] = [];
    for (let j = 0; j < LAND_NUMBER; j++) g[i][j] = makeStatus();
  }
  return g;
}

export const state = {
  gameMode: GameMode.Entrance,
  floor: 0,
  turn: 200,
  sumDamage: 0,
  monsterNumber: 10,

  player: makeStatus(),

  monsters: Array.from({length:1000}, makeStatus),
  landsquare: makeLandGrid(),
  randomPosition: makeStatus(),

  abilityHp: new Array(7).fill(0),
  words: new Array(12).fill(''),
  oword: Array.from({length:4}, makeStatus),

  // loaded map data
  mapPixels: null,   // Uint8ClampedArray (RGBA)
  mapW: 0, mapH: 0,
  colorRef: [],      // colorRef[1..6] = {r,g,b}

  // sprite OffscreenCanvases
  sprites: { player:null, box:null, monster:null, room:null, wall:null },

  tileLayerDirty: true,
  tileLayerCanvas: null,

  autoPlay: false,

  // DLC: boss floors + persistent achievement records (localStorage 'dungeon_records')
  bossIdx: -1,
  records: { deepestFloor: 0, clears: 0, totalKills: 0, totalBoxes: 0, bossKills: 0, runs: 0 },
};
