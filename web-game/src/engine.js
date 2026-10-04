// ゲームロジック本体。VB6 の BasicModule / LandModule / AbilityModule /
// MonsterModule / WordModule / Form1 のロジック部分を移植したもの。
// DOM には依存しない（描画は renderer.js、音は audio.js が g を読んで行う）。

import {
  ENTRANCE, DUNGEON, HOWTO_PLAY, OPTIONS, GAME_OVER, GAME_CLEAR, MUSEUM, BLESSING,
  BLUE_BOX, RED_BOX, YELLOW_BOX, GREEN_BOX, PURPLE_BOX, STAIR, WALL, ROOM, ENEMY,
  LAND_NUMBER, NO_ABILITY, WALL_BREAK, SLOW, BOX_ATTACK, TURN_CONST, STEALTH,
  MONSTER_MAX, DIFFICULTY_NAMES, SCREEN, FONT_SIZE,
  DIR_NONE, DIR_UP, DIR_DOWN, DIR_RIGHT, DIR_LEFT, DIR_IDLE,
  MUSIC_MENU, MUSIC_CLEAR, MUSIC_GAMEOVER, MUSIC_DUNGEON,
} from './constants.js';
import { cLng, vbInt, rnd } from './vb.js';
import { getPixel } from './bmp.js';
import { generateFloor } from './mapgen.js';

const LIMIT_9 = 10 ** 9 - 1;
const LIMIT_7 = 10 ** 7 - 1;
const LIMIT_6 = 10 ** 6 - 1;
const LIMIT_4 = 10 ** 4 - 1;

function makeStatus() {
  return {
    x: 0, y: 0, ox: 0, oy: 0, // Left/Top, Oleft/Otop（マス単位）
    alive: false,
    hp: 0, maxHp: 0, exp: 0, atk: 0, ability: 0, def: 0, level: 0,
    condition: 0, direction: DIR_NONE,
    name: '', explanation: '',
  };
}

// 1 回の冒険の集計（リザルト画面と記録用。ゲームの進行には影響しない）
function makeStats() {
  return {
    boxes: [0, 0, 0, 0, 0], walls: 0, kills: 0, maxDamage: 0, damageTaken: 0,
    commands: 0, revives: 0, blessings: 0, turns: 0, auto: false,
  };
}

// ゲーム状態（VB6 の Public 変数群）
export const g = {
  mode: ENTRANCE,
  floor: 0,
  turn: 0,
  sumDamage: 0,
  monsterNumber: 0,
  player: makeStatus(),
  monsters: [],
  land: [], // land[x][y]
  abilityHp: [0, 0, 0, 0, 0, 0, 0],
  words: new Array(12).fill(''),
  owords: [0, 1, 2, 3].map(() => ({ explanation: '', hp: 0 })),
  foreColor: '#ffffff',
  backColor: '#000000',
  music: MUSIC_MENU, // 鳴らすべき曲
  musicLoop: true,
  musicToken: 0, // 同じ曲を頭から鳴らし直したいときに増やす
  demo: [], // How to play 画面の見本スプライト
  mapColors: [], // ColorSet で読む基準色
  saveExists: false,
  // ---- ここから下はブラウザ版独自
  events: [], // 演出用イベント（renderer / sfx が毎フレーム取り出す）
  stats: makeStats(),
  records: { best: 0, runs: [] }, // 冒険の記録（上位 10 件）
  result: null, // 直近の冒険の結果（GameOver / GameClear 画面用）
  arrange: true, // アレンジルール（10 階ごとの祝福）。false なら VB6 版そのまま
  blessing: null, // 祝福の選択画面の状態 { options, cursor, lock }
};

// 演出用イベントを積む。取り出す側が居ない（ヘッドレス）ときは溜めすぎない。
function emit(t, data) {
  if (g.events.length < 256) g.events.push({ t, ...data });
}

let mapImage = null;
let storage = {
  data: null,
  load() { return this.data; },
  save(d) { this.data = d; },
};

// mapImg: decodeBmp() の結果（Map40.bmp）, store: {load(), save(obj)}
export function initGame(mapImg, store) {
  mapImage = mapImg;
  if (store) storage = store;
  const rec = storage.loadRecords?.();
  g.records = { best: rec?.best ?? 0, runs: Array.isArray(rec?.runs) ? rec.runs : [] };
  g.arrange = rec?.arrange !== false;
  g.events.length = 0;
  // Form_Load
  g.mode = ENTRANCE;
  musicSet();
  monsterSet();
  landSet();
  wordSet();
}

const inBounds = (x, y) => x >= 0 && y >= 0 && x < LAND_NUMBER && y < LAND_NUMBER;
const isEdge = (x, y) => x === 0 || y === 0 || x === LAND_NUMBER - 1 || y === LAND_NUMBER - 1;

function tileName(cond) {
  switch (cond) {
    case BLUE_BOX: return '青箱';
    case RED_BOX: return '赤箱';
    case YELLOW_BOX: return '黄箱';
    case GREEN_BOX: return '緑箱';
    case PURPLE_BOX: return '紫箱';
    case WALL: return '壁';
    default: return '';
  }
}

function playMusic(track, loop) {
  g.music = track;
  g.musicLoop = loop;
  g.musicToken++;
}

// ---------------------------------------------------------------- 入力

// Form_KeyDown。key は 'Up' 'Down' 'Left' 'Right' 'Enter' 'A'..'Z'
export function keyDown(key) {
  const p = g.player;
  switch (g.mode) {
    case ENTRANCE:
      switch (key) {
        case 'Z': g.mode = HOWTO_PLAY; landSet(); monsterSet(); break;
        case 'C': g.mode = OPTIONS; break;
        case 'D': g.mode = MUSEUM; landSet(); break;
        case 'Enter': g.mode = DUNGEON; landSet(); monsterSet(); floorSet(); break;
      }
      break;
    case DUNGEON:
      switch (key) {
        case 'Up': p.direction *= DIR_UP; break;
        case 'Down': p.direction *= DIR_DOWN; break;
        case 'Right': p.direction *= DIR_RIGHT; break;
        case 'Left': p.direction *= DIR_LEFT; break;
        case 'Z': case 'X': case 'C': case 'D': case 'S': case 'Enter': case 'A':
          abilityEffect(key);
          break;
      }
      break;
    case HOWTO_PLAY:
      switch (key) {
        case 'Right': p.direction *= DIR_RIGHT; break;
        case 'Left': p.direction *= DIR_LEFT; break;
        case 'X': backToEntrance(); break;
      }
      break;
    case OPTIONS:
      switch (key) {
        case 'Right': p.direction *= DIR_RIGHT; break;
        case 'Left': p.direction *= DIR_LEFT; break;
        case 'Up': case 'Down': g.arrange = !g.arrange; saveRecords(); emit('select'); break;
        case 'X': backToEntrance(); break;
        case 'Enter': g.mode = DUNGEON; landSet(); monsterSet(); floorSet(); break;
      }
      break;
    case BLESSING:
      blessingKey(key);
      break;
    case GAME_OVER:
    case GAME_CLEAR:
      if (key === 'X') { musicSet(); backToEntrance(); }
      break;
    case MUSEUM:
      switch (key) {
        case 'X': backToEntrance(); break;
        // port: セーブが無いときのロードは VB6 では型エラーになるので無視する
        case 'L': if (g.saveExists) abilityEffect('L'); break;
      }
      break;
  }
}

function backToEntrance() {
  g.mode = ENTRANCE;
  landSet();
  wordSet();
}

// ---------------------------------------------------------------- 1 フレーム (Timer1_Timer, 50ms)

export function tick() {
  switch (g.mode) {
    case ENTRANCE:
    case OPTIONS:
    case GAME_OVER:
    case GAME_CLEAR:
    case MUSEUM:
      wordSet();
      musicCheck();
      movement();
      break;
    case DUNGEON:
      if (g.player.alive) movement();
      // port: 移動中にセーブ/クリアで画面が変わった場合は、そのフレームの判定を行わない
      if (g.mode === DUNGEON) {
        statusCheck(false);
        wordSet();
        musicCheck();
      }
      break;
    case HOWTO_PLAY:
      wordSet();
      movement();
      break;
    case BLESSING:
      if (g.blessing.lock > 0) g.blessing.lock--;
      break;
  }
}

// ---------------------------------------------------------------- BasicModule

function moveMonsterToward(m) {
  const p = g.player;
  if (m.y > p.y) { m.y--; m.direction *= DIR_UP; }
  if (m.y < p.y) { m.y++; m.direction *= DIR_DOWN; }
  if (m.x < p.x) { m.x++; m.direction *= DIR_RIGHT; }
  if (m.x > p.x) { m.x--; m.direction *= DIR_LEFT; }
}

function movement() {
  const p = g.player;
  switch (g.mode) {
    case ENTRANCE:
      p.direction = DIR_NONE;
      break;

    case DUNGEON: {
      if (p.direction <= 1) break;
      if (p.direction % 2 === 0 && p.y > 0) p.y--;
      if (p.direction % 3 === 0 && p.y < LAND_NUMBER - 1) p.y++;
      if (p.direction % 5 === 0 && p.x < LAND_NUMBER - 1) p.x++;
      if (p.direction % 7 === 0 && p.x > 0) p.x--;

      positionCheckPlayer();
      // port: 階段でセーブして画面が変わった/1000F に着いた場合はモンスターを動かさない
      if (g.mode !== DUNGEON) return;

      if (p.condition === SLOW && g.turn % 2 === 0) {
        p.direction = DIR_IDLE;
      } else {
        p.direction = DIR_NONE;
      }
      g.turn--;
      g.stats.turns++;

      for (let i = 0; i < g.monsterNumber; i++) {
        const m = g.monsters[i];
        if (!m.alive || p.condition === STEALTH) continue;
        if (m.ability !== SLOW || g.turn % 2 === 0) {
          moveMonsterToward(m);
          positionCheckMonster(i);
          m.direction = DIR_NONE;
        }
      }

      if (g.sumDamage > 0) {
        g.words[3] = `プレイヤーは${g.sumDamage}ダメージを受けた`;
        g.stats.damageTaken += g.sumDamage;
        emit('hurt', { dmg: g.sumDamage, ratio: g.sumDamage / p.maxHp });
        g.sumDamage = 0;
      }
      break;
    }

    case HOWTO_PLAY: {
      const pages = HOWTO_PAGES.length;
      if (p.direction % 5 === 0) p.condition = (p.condition + 1) % pages;
      else if (p.direction % 7 === 0) p.condition = p.condition === 0 ? pages - 1 : p.condition - 1;
      p.direction = DIR_NONE;
      setupDemo(p.condition);
      break;
    }

    case OPTIONS:
      if (p.direction % 5 === 0) p.ability = (p.ability + 1) % 5;
      else if (p.direction % 7 === 0) p.ability = p.ability === 0 ? 4 : p.ability - 1;
      p.direction = DIR_NONE;
      break;

    case MUSEUM:
      // port: VB6 では左右キーで Condition が変わるがページは 1 枚しかないため何もしない
      p.direction = DIR_NONE;
      break;
  }
}

// How to play の見本スプライト（ピクセル座標は VB6 の配置をそのまま使う）
function setupDemo(page) {
  const half = SCREEN / 2;
  const top = FONT_SIZE * 10.5; // Player.Top
  g.demo = [];
  if (page === 1) {
    for (let i = 0; i < 5; i++) {
      g.demo.push({ sheet: 'player', row: i, left: half + 20 * i, top });
      g.demo.push({ sheet: 'monster', row: i, left: half + 20 * i, top: top + FONT_SIZE * 2 });
    }
    g.demo.push({ sheet: 'box', row: STAIR, left: half, top: top + FONT_SIZE * 4 });
    for (let i = 0; i < 10; i++) {
      g.demo.push({ sheet: 'wall', row: i, left: half + 20 * (i - 4), top: top + FONT_SIZE * 14 });
      g.demo.push({ sheet: 'room', row: i, left: half + 20 * (i - 4), top: top + FONT_SIZE * 16 });
    }
  } else if (page === 2) {
    // port: VB6 では 3〜6 ページ全てで見出しに重なって表示されていたため、
    // 箱の説明ページ (3/6) の空行にだけ表示する
    for (let i = 0; i < 5; i++) {
      g.demo.push({ sheet: 'box', row: i, left: half + 20 * (i - 2), top: FONT_SIZE * 2 * 4 + 9 });
    }
  }
}

export function statusCheck(levelup) {
  const p = g.player;

  if (p.exp >= p.level ** 3 && p.level < 999 && p.alive) {
    const dHp = vbInt(5 + rnd() * 3);
    p.level++;
    p.maxHp += dHp;
    p.hp += dHp;
    p.atk = cLng(p.atk * 1.1 + 1);
    p.def++;
    emit('level', { level: p.level });
  }

  if (p.atk > LIMIT_9) p.atk = LIMIT_9;
  if (p.def > LIMIT_9) p.def = LIMIT_9;
  if (p.def === 0) p.def = 1;
  if (g.turn > LIMIT_4) g.turn = LIMIT_4;
  if (p.exp > LIMIT_9) p.exp = LIMIT_9;
  if (p.maxHp > LIMIT_7) p.maxHp = LIMIT_7;

  if (p.hp > p.maxHp) p.hp = p.maxHp;
  else if (p.hp > p.maxHp / 4) g.foreColor = '#ffffff';
  else if (p.hp > 0) g.foreColor = 'rgb(255,127,39)';

  if (p.hp <= 0 || g.turn <= 0) {
    if (g.abilityHp[5] === 0) {
      p.direction = DIR_NONE;
      p.alive = false;
      if (g.mode === DUNGEON) finishRun(g.turn <= 0 && p.hp > 0 ? 'turn' : 'hp');
      p.hp = 0;
      g.mode = GAME_OVER;
      landSet();
      wordSet();
    } else {
      p.alive = false;
      abilityEffect('D'); // 復活の珠
    }
  }

  for (let i = 0; i < g.monsters.length; i++) {
    const m = g.monsters[i];
    if (m.level < 999 && levelup) {
      m.level++;
      m.maxHp = cLng(m.maxHp * 1.1 + 1);
      m.hp = m.maxHp;
      m.atk = cLng(m.atk * 1.2 + 1);
      m.def++;
      m.exp += 10;
      if (i === 0) g.words[0] = `ヘキサスライム達はLevel${m.level}になった。`;
    }
    if (m.atk > LIMIT_9) m.atk = LIMIT_9;
    if (m.def > LIMIT_9) m.def = LIMIT_9;
    if (m.def === 0) m.def = 1;
    if (m.exp > LIMIT_4) m.exp = LIMIT_4;
    if (m.hp > m.maxHp) m.hp = m.maxHp;
    if (m.maxHp > LIMIT_6) m.maxHp = LIMIT_6;
  }

  if (g.mode === DUNGEON || g.mode === MUSEUM || g.mode === GAME_OVER) {
    for (const col of g.land) {
      for (const t of col) {
        if (t.def > LIMIT_9) t.def = LIMIT_9;
        if (t.def === 0) t.def = 1;
        if (t.hp > t.maxHp) t.hp = t.maxHp;
        if (t.maxHp > LIMIT_6) t.maxHp = LIMIT_6;
      }
    }
  }
}

export function floorSet() {
  const p = g.player;
  g.floor++;

  switch (g.floor) {
    case 1:
      g.words.fill('');
      Object.assign(p, {
        hp: 20, maxHp: 20, level: 1, exp: 0, x: 0, y: 0,
        atk: 10, def: 3, direction: DIR_NONE, condition: 0,
      });
      g.turn = 200;
      g.monsterNumber = 10;
      g.stats = makeStats();
      g.result = null;
      break;
    case 20: case 40: case 60: case 80: case 100:
      g.monsterNumber += 10;
      break;
    case 200: case 300: case 400: case 500: case 600: case 700: case 800:
      g.monsterNumber += 100;
      break;
    case 900:
      g.monsterNumber = 1000;
      break;
    case 1000:
      p.direction = DIR_NONE;
      p.alive = false;
      finishRun('clear');
      g.mode = GAME_CLEAR;
      landSet();
      wordSet();
      return; // port: VB6 はこの後も裏でマップを作るが、画面には出ないので省略
  }

  // arrange: フロアをランダムに作る（VB6 版は Map40.bmp の 64 枚から選ぶ）
  const gen = g.arrange ? generateFloor(g.floor) : null;
  mapSet(gen);

  // VB6 は「ランダムなマスを引いて床ならそこに置く」を繰り返す。
  // 空き床から一様に選ぶのと同じ分布なので、候補リストから引く形にした。
  const rooms = [];
  for (let x = 0; x < LAND_NUMBER; x++) {
    for (let y = 0; y < LAND_NUMBER; y++) {
      if (g.land[x][y].condition === ROOM) rooms.push(g.land[x][y]);
    }
  }
  if (gen) placeGenerated(gen, rooms);
  const take = () => {
    if (rooms.length === 0) return null;
    const k = vbInt(rnd() * rooms.length);
    const t = rooms[k];
    rooms[k] = rooms[rooms.length - 1];
    rooms.pop();
    return t;
  };

  // プレイヤー
  const pt = gen ? null : take();
  if (pt) {
    p.x = p.ox = pt.x;
    p.y = p.oy = pt.y;
    pt.alive = false;
  }
  p.alive = true;
  p.condition = NO_ABILITY;

  // 階段・紫箱
  const st = gen ? null : take();
  if (st) st.condition = STAIR;
  const pb = gen ? null : take();
  if (pb) pb.condition = PURPLE_BOX;

  // モンスター
  for (let i = 0; i < g.monsters.length; i++) {
    const m = g.monsters[i];
    m.condition = 0;
    const t = i < g.monsterNumber ? take() : null;
    if (t) {
      m.x = m.ox = t.x;
      m.y = m.oy = t.y;
      m.alive = true;
      m.hp = m.maxHp;
      t.condition = ENEMY;
    } else {
      m.alive = false;
    }
  }

  if (g.floor >= 2) statusCheck(true);
  emit('floor', { floor: g.floor });
  if (gen?.text) g.words[1] = gen.text;
  // arrange: 10 階を越えるごとに祝福を 1 つ選ぶ
  if (g.arrange && g.floor > 1 && g.floor % 10 === 1) blessingSet();
}

// arrange: ランダム生成したフロアでは、プレイヤーは決められたスタート地点に立つ
// （階段と紫箱は generateFloor() が置き済み）。モンスターはスタート地点の近くには置かない。
function placeGenerated(gen, rooms) {
  const p = g.player;
  p.x = p.ox = gen.x;
  p.y = p.oy = gen.y;
  g.land[gen.x][gen.y].alive = false;
  const far = rooms.filter((t) => Math.max(Math.abs(t.x - gen.x), Math.abs(t.y - gen.y)) > gen.safe);
  // 遠いマスだけでは置ききれないほどモンスターが多い階は、スタート地点以外のどこにでも置く
  const keep = far.length >= g.monsterNumber ? far : rooms.filter((t) => t.x !== gen.x || t.y !== gen.y);
  rooms.length = 0;
  rooms.push(...keep);
}

// ---------------------------------------------------------------- 祝福（アレンジルール）

const BLESSINGS = [
  {
    name: '剛力の祝福', desc: '攻撃力が\n1.5倍になる',
    apply(p) { p.atk = cLng(p.atk * 1.5) + 2; },
  },
  {
    name: '鉄壁の祝福', desc: '守備力が\n1.5倍になる',
    apply(p) { p.def = cLng(p.def * 1.5) + 2; },
  },
  {
    name: '生命の祝福', desc: '最大Hpが1.5倍になり\nHpが全快する',
    apply(p) { p.maxHp = cLng(p.maxHp * 1.5) + 10; p.hp = p.maxHp; },
  },
  {
    name: '時の祝福', desc: 'ターン数が\n200増える',
    apply() { g.turn += 200; },
  },
  {
    name: '紫の祝福', desc: '5つのコマンドの\n使用回数が1増える',
    apply() { for (let k = 0; k <= 4; k++) g.abilityHp[k]++; },
  },
  {
    name: '再生の祝福', desc: '復活の珠が\n2個手に入る',
    apply() { g.abilityHp[5] += 2; },
  },
  {
    name: '衰弱の祝福', desc: '全てのモンスターの\n攻撃力が半分になる',
    apply() { for (const m of g.monsters) m.atk = cLng(m.atk / 2); },
  },
  {
    name: '豊穣の祝福', desc: 'このフロアの箱が\n全て緑色になる',
    apply() {
      for (const col of g.land) for (const b of col) if (b.condition <= GREEN_BOX) b.condition = GREEN_BOX;
    },
  },
];

export function blessingInfo(id) {
  return BLESSINGS[id];
}

function blessingSet() {
  const pool = BLESSINGS.map((_, i) => i);
  const options = [];
  for (let k = 0; k < 3; k++) options.push(pool.splice(vbInt(rnd() * pool.length), 1)[0]);
  // lock: 降りてきた勢いのキー入力で誤って決定しないよう、少しの間入力を受け付けない
  g.blessing = { options, cursor: 1, lock: 8, floor: g.floor - 1 };
  g.mode = BLESSING;
  emit('bless');
}

function blessingKey(key) {
  const b = g.blessing;
  if (b.lock > 0) return;
  switch (key) {
    case 'Left': b.cursor = (b.cursor + 2) % 3; emit('select'); break;
    case 'Right': b.cursor = (b.cursor + 1) % 3; emit('select'); break;
    case 'Enter': case 'Z': {
      const chosen = BLESSINGS[b.options[b.cursor]];
      chosen.apply(g.player);
      g.stats.blessings++;
      g.mode = DUNGEON;
      g.player.direction = DIR_NONE;
      g.words[1] = `${chosen.name}を受けた。`;
      statusCheck(false);
      emit('blessed', { name: chosen.name });
      break;
    }
  }
}

// ---------------------------------------------------------------- 記録・称号

const TITLES = [
  [1000, '果報者'], [500, '箱に愛されし者'], [300, '深淵の住人'], [200, '深淵を覗く者'],
  [100, '百階の覇者'], [75, '不思議の探求者'], [50, '迷宮の常連'], [30, '箱の目利き'],
  [20, 'スライム狩り'], [10, '箱開け見習い'], [5, '駆け出しの探索者'], [0, '箱を知らぬ者'],
];

export function titleFor(floor) {
  return TITLES.find(([f]) => floor >= f)[1];
}

function saveRecords() {
  storage.saveRecords?.({ best: g.records.best, runs: g.records.runs, arrange: g.arrange });
}

// 冒険の終わり（cause: 'hp' | 'turn' | 'clear'）。オートプレイの結果は自己ベストに数えない。
function finishRun(cause) {
  const p = g.player;
  const r = g.records;
  const run = {
    floor: g.floor, level: p.level, cause, arrange: g.arrange,
    difficulty: p.ability, auto: g.stats.auto, date: Date.now(),
  };
  const newBest = !run.auto && run.floor > r.best;
  if (newBest) r.best = run.floor;
  r.runs.push(run);
  r.runs.sort((a, b) => b.floor - a.floor || b.date - a.date);
  if (r.runs.length > 10) r.runs.length = 10;
  g.result = { ...run, newBest, rank: r.runs.indexOf(run) + 1 };
  saveRecords();
  emit(cause === 'clear' ? 'clear' : 'gameover', { newBest });
}

function positionCheckPlayer() {
  const p = g.player;
  const t = g.land[p.x][p.y];

  if (t.condition === STAIR) {
    if (g.abilityHp[6] === 0) floorSet();
    else abilityEffect('S');
    if (g.mode !== DUNGEON) return;
  } else if (t.condition <= PURPLE_BOX) {
    battlePlayerTile(p.x, p.y);
    p.x = p.ox;
    p.y = p.oy;
  } else if (t.condition === WALL) {
    if (p.condition === WALL_BREAK && !isEdge(p.x, p.y)) battlePlayerTile(p.x, p.y);
    p.x = p.ox;
    p.y = p.oy;
  }

  for (let i = 0; i < g.monsters.length; i++) {
    const m = g.monsters[i];
    if (m.alive && m.x === p.x && m.y === p.y) {
      battlePlayerMonster(i);
      p.x = p.ox;
      p.y = p.oy;
    }
  }

  p.ox = p.x;
  p.oy = p.y;
}

// 進めなかったときの回り込み候補（Direction ごとに VB6 と同じ順で試す）
const DETOURS = {
  2: [[-1, -1], [1, -1]],
  3: [[1, 1], [-1, 1]],
  5: [[1, -1], [1, 1]],
  7: [[-1, 1], [-1, -1]],
  10: [[0, -1], [1, 0]],
  15: [[1, 0], [0, 1]],
  21: [[0, 1], [-1, 0]],
  14: [[-1, 0], [0, -1]],
};

function positionCheckMonster(i) {
  const m = g.monsters[i];
  const p = g.player;
  if (!m.alive) return;

  if (m.x === p.x && m.y === p.y && p.alive) {
    m.x = m.ox;
    m.y = m.oy;
    battleMonsterTarget(i);
  } else if (g.land[m.x][m.y].condition !== ROOM) {
    const c = g.land[m.x][m.y].condition;
    if (c === WALL && m.ability === WALL_BREAK && !isEdge(m.x, m.y)) {
      battleMonsterTile(i, m.x, m.y);
    } else if (c <= PURPLE_BOX && m.ability === BOX_ATTACK) {
      battleMonsterTile(i, m.x, m.y);
    }

    const detours = DETOURS[m.direction];
    if (detours) {
      m.x = m.ox;
      m.y = m.oy;
      for (const [dx, dy] of detours) {
        const nx = m.ox + dx;
        const ny = m.oy + dy;
        // port: 一番外側の壁が消えたとき VB6 は配列外参照で落ちるので範囲チェックを追加
        if (inBounds(nx, ny) && g.land[nx][ny].condition === ROOM) {
          m.x = nx;
          m.y = ny;
          break;
        }
      }
    }
  }

  g.land[m.ox][m.oy].condition = ROOM;
  g.land[m.x][m.y].condition = ENEMY;
  m.ox = m.x;
  m.oy = m.y;
}

function calcDamage(atk, def) {
  const d = vbInt((atk * (rnd() * 0.2 + 0.9)) / def) + 1;
  return d > LIMIT_7 ? LIMIT_7 : d;
}

// BattleCheck(PlayerConst, i, j)  j < LandNumber
function battlePlayerTile(x, y) {
  const t = g.land[x][y];
  const damage = calcDamage(g.player.atk, t.def);
  t.hp -= damage;
  g.words[2] = `${tileName(t.condition)}は${damage}ダメージを受けた`;
  if (damage > g.stats.maxDamage) g.stats.maxDamage = damage;
  emit('hit', { x, y, dmg: damage, tile: true });
  if (t.hp <= 0) {
    t.hp = 0;
    t.ability = vbInt(rnd() * 15);
    if (g.player.condition !== TURN_CONST) g.turn += 10;
    boxEffect(x, y);
  }
}

// BattleCheck(PlayerConst, i, LandNumber)
function battlePlayerMonster(i) {
  const m = g.monsters[i];
  if (!m.alive) return;
  const damage = calcDamage(g.player.atk, m.def);
  m.hp -= damage;
  g.words[2] = `${m.name}は${damage}ダメージを受けた`;
  if (damage > g.stats.maxDamage) g.stats.maxDamage = damage;
  emit('hit', { x: m.x, y: m.y, dmg: damage });
  if (m.hp <= 0) {
    m.hp = 0;
    m.alive = false;
    g.land[m.x][m.y].condition = ROOM;
    g.player.exp += m.exp;
    g.stats.kills++;
    emit('kill', { x: m.x, y: m.y });
  }
}

// BattleCheck(i, x, y)  モンスター → 地形
function battleMonsterTile(i, x, y) {
  const t = g.land[x][y];
  if (t.condition === STAIR) return;
  const damage = calcDamage(g.monsters[i].atk, t.def);
  t.hp -= damage;
  g.words[2] = `${tileName(t.condition)}は${damage}ダメージを受けた`;
  if (t.hp <= 0) {
    t.hp = 0;
    t.condition = ROOM;
  }
}

// BattleCheck(i, PlayerConst, LandNumber)  モンスター → プレイヤー
function battleMonsterTarget(i) {
  const damage = calcDamage(g.monsters[i].atk, g.player.def);
  g.player.hp -= damage;
  g.sumDamage += damage;
}

function musicCheck() {
  switch (g.mode) {
    case ENTRANCE: case HOWTO_PLAY: case OPTIONS: case MUSEUM:
      if (g.music !== MUSIC_MENU) playMusic(MUSIC_MENU, true);
      break;
    case DUNGEON:
      if (g.music !== MUSIC_DUNGEON) playMusic(MUSIC_DUNGEON, true);
      break;
  }
}

// Form1.MusicSet — GameOver/GameClear から戻る時に曲の状態をリセット
function musicSet() {
  if (g.mode === GAME_OVER || g.mode === GAME_CLEAR) g.music = -1;
}

// ---------------------------------------------------------------- Form1: ColorSet / MapSet

function colorSet() {
  g.mapColors = [];
  if (!mapImage) return;
  for (let i = 1; i <= 6; i++) g.mapColors.push(getPixel(mapImage, i, 1));
}

const COLOR_TO_CONDITION = [BLUE_BOX, RED_BOX, YELLOW_BOX, GREEN_BOX, WALL, ROOM];

function matchColor(z) {
  let best = -1;
  let bestDist = Infinity;
  for (let k = 0; k < g.mapColors.length; k++) {
    const c = g.mapColors[k];
    if (c === z) return COLOR_TO_CONDITION[k];
    const d =
      Math.abs(((c >> 16) & 255) - ((z >> 16) & 255)) +
      Math.abs(((c >> 8) & 255) - ((z >> 8) & 255)) +
      Math.abs((c & 255) - (z & 255));
    if (d < bestDist) { bestDist = d; best = k; }
  }
  // port: 一致しない色は VB6 では前の階の状態が残る。最も近い色として扱う。
  return best >= 0 ? COLOR_TO_CONDITION[best] : ROOM;
}

function mapSet(gen) {
  const m0 = g.monsters[0];
  let ox = 0;
  let oy = 0;
  if (mapImage && !gen) {
    ox = vbInt(rnd() * vbInt(mapImage.width / LAND_NUMBER)) * LAND_NUMBER;
    oy = vbInt(rnd() * vbInt(mapImage.height / LAND_NUMBER)) * LAND_NUMBER;
  }
  for (let x = 0; x < LAND_NUMBER; x++) {
    for (let y = 0; y < LAND_NUMBER; y++) {
      const t = g.land[x][y];
      if (gen) {
        t.condition = gen.cond[x * LAND_NUMBER + y];
      } else if (mapImage) {
        t.condition = matchColor(getPixel(mapImage, ox + x, oy + y));
      } else {
        t.condition = isEdge(x, y) ? WALL : ROOM;
      }
      t.maxHp = m0.maxHp;
      t.hp = t.maxHp;
      t.def = m0.def;
      t.alive = true;
    }
  }
}

// ---------------------------------------------------------------- MonsterModule

function monsterSet() {
  if (g.mode === ENTRANCE || g.mode === DUNGEON) {
    while (g.monsters.length < MONSTER_MAX) g.monsters.push(makeStatus());
    g.monsters.forEach((m, i) => {
      Object.assign(m, {
        name: `ヘキサスライム${i}`, explanation: '',
        hp: 10, maxHp: 10, level: 1, exp: 10, ability: NO_ABILITY,
        atk: 10, def: 1, direction: DIR_NONE,
      });
    });
  }
  // HowtoPlay の見本モンスターは setupDemo() で扱う
}

// ---------------------------------------------------------------- LandModule

function landSet() {
  const p = g.player;
  switch (g.mode) {
    case ENTRANCE:
      g.abilityHp = [0, 0, 0, 0, 0, 0, 0];
      Object.assign(p, { alive: false, level: 1, maxHp: 20, hp: 20, condition: 0, direction: DIR_NONE });
      g.land = [];
      for (let x = 0; x < LAND_NUMBER; x++) {
        const col = [];
        for (let y = 0; y < LAND_NUMBER; y++) {
          const t = makeStatus();
          t.x = x;
          t.y = y;
          t.condition = ROOM;
          t.alive = true;
          col.push(t);
        }
        g.land.push(col);
      }
      colorSet();
      g.backColor = '#000000';
      g.foreColor = '#ffffff';
      g.floor = 0;
      g.demo = [];
      break;

    case DUNGEON:
      g.words.fill('');
      p.hp = p.maxHp;
      p.alive = true;
      break;

    case HOWTO_PLAY:
      setupDemo(p.condition);
      break;

    case GAME_OVER:
      g.foreColor = '#ffffff';
      playMusic(MUSIC_GAMEOVER, false);
      break;

    case GAME_CLEAR:
      playMusic(MUSIC_CLEAR, false);
      break;

    case MUSEUM: {
      const d = storage.load();
      g.saveExists = !!d;
      if (!d) break;
      Object.assign(p, {
        hp: d.player.hp, maxHp: d.player.maxHp, atk: d.player.atk, def: d.player.def,
        exp: d.player.exp, level: d.player.level, ability: d.player.ability,
      });
      for (let i = 0; i <= 5; i++) g.abilityHp[i] = d.abilityHp[i];
      g.floor = d.floor;
      g.turn = d.turn;
      g.monsterNumber = d.monsterNumber;
      statusCheck(false);
      break;
    }
  }
}

function boxEffect(x, y) {
  const t = g.land[x][y];
  const box = t.condition;
  const p = g.player;
  const ms = g.monsters;
  const eachBox = (pred, fn) => {
    for (const col of g.land) for (const b of col) if (pred(b.condition)) fn(b);
  };
  const killAll = () => {
    for (const m of ms) {
      if (m.alive) {
        m.alive = false;
        m.hp = 0;
        m.condition = 0;
        g.land[m.x][m.y].condition = ROOM;
      }
    }
  };
  const isBox = (c) => c <= PURPLE_BOX;
  const isColorBox = (c) => c <= GREEN_BOX;

  switch (t.condition) {
    case BLUE_BOX:
      p.hp = cLng(p.hp + p.maxHp / 10);
      for (const m of ms) m.ability = NO_ABILITY;
      t.explanation = 'HPが少し回復し、モンスターの状態異常が解除された。';
      break;

    case RED_BOX:
      switch (t.ability) {
        case 0: p.hp = 1; t.explanation = 'Hpが1になってしまった'; break;
        case 1: p.atk = cLng(p.atk * 0.9); t.explanation = '攻撃力が下がった'; break;
        case 2: p.condition = SLOW; t.explanation = 'プレイヤーはこのフロアにいる間、動きが鈍くなった。'; break;
        case 3: p.condition = TURN_CONST; t.explanation = 'このフロアにいる間、ターン数が増えなくなった。'; break;
        case 4: p.maxHp = cLng(p.maxHp * 0.9); t.explanation = '最大Hpが下がった'; break;
        case 5: p.hp = vbInt(p.hp / 2) + 1; t.explanation = 'Hpが半分になった'; break;
        case 6:
          for (const m of ms) if (m.alive) { m.hp = m.maxHp; m.condition = 0; }
          t.explanation = 'モンスターが全快した。';
          break;
        case 7:
          for (const m of ms) m.atk = cLng(m.atk * 1.1);
          t.explanation = '全てのモンスターの攻撃力が少し上がった';
          break;
        case 8:
          statusCheck(true);
          t.explanation = '全てのモンスターのレベルが上った';
          break;
        case 9:
          eachBox(isBox, (b) => { b.def *= 2; });
          t.explanation = 'このフロアの全ての箱の守備力が2倍になった。';
          break;
        case 10:
          for (const m of ms) m.def = cLng(m.def * 1.1);
          t.explanation = '全てのモンスターの守備力が少し上がった';
          break;
        case 11:
          eachBox(isColorBox, (b) => { b.condition = RED_BOX; });
          t.explanation = '全ての箱が赤色になった';
          break;
        case 12:
          for (const m of ms) m.maxHp = cLng(m.maxHp * 1.1);
          t.explanation = '全てのモンスターの最大Hpが少し上がった';
          break;
        case 13:
          for (const m of ms) m.exp = cLng(m.exp * 0.9);
          t.explanation = '全てのモンスターの経験値が少し下がった';
          break;
        case 14:
          eachBox(isBox, (b) => { b.hp *= 2; b.maxHp *= 2; });
          t.explanation = 'このフロアの全ての箱のHpが2倍になった。';
          break;
      }
      break;

    case YELLOW_BOX:
      switch (t.ability) {
        case 0:
          for (const m of ms) m.ability = STEALTH;
          t.explanation = 'モンスターが透明になった。';
          break;
        case 1:
          for (let k = 0; k < g.abilityHp.length - 1; k++) g.abilityHp[k]++;
          t.explanation = '全てのコマンドの使用回数が1増えた。';
          break;
        case 2: p.condition = TURN_CONST; t.explanation = 'このフロアにいる間、ターン数が増えなくなった。'; break;
        case 3: killAll(); t.explanation = 'モンスターが全滅した。'; break;
        case 4: p.def *= 2; t.explanation = '守備力が2倍になった'; break;
        case 5: p.atk *= 2; t.explanation = '攻撃力が2倍になった'; break;
        case 6: g.turn = 1000; t.explanation = 'ターン数が残り1000になった。'; break;
        case 7: p.atk = cLng(p.atk / 2); t.explanation = '攻撃力が半分になった'; break;
        case 8:
          for (const m of ms) m.ability = WALL_BREAK;
          t.explanation = 'モンスターが壁を掘れるようになった。(一番外側を除く)';
          break;
        case 9:
          for (let k = 0; k < g.abilityHp.length - 1; k++) g.abilityHp[k] = 5;
          t.explanation = '全てのコマンドの使用回数が5になった。';
          break;
        case 10:
          eachBox(isColorBox, (b) => { b.condition = ROOM; });
          t.explanation = '全ての箱が消滅した。';
          break;
        case 11:
          eachBox(isColorBox, (b) => { b.condition = BLUE_BOX; });
          t.explanation = '全ての箱が青色になった';
          break;
        case 12:
          eachBox(isColorBox, (b) => { b.condition = vbInt(rnd() * 4); });
          t.explanation = '全ての箱がランダムに変化した。';
          break;
        case 13:
          eachBox(isColorBox, (b) => { b.condition = YELLOW_BOX; });
          t.explanation = '全ての箱が黄色になった';
          break;
        case 14:
          for (const m of ms) m.ability = BOX_ATTACK;
          t.explanation = 'モンスターが箱を壊せるようになった。';
          break;
      }
      break;

    case GREEN_BOX:
      switch (t.ability) {
        case 0: p.hp = p.maxHp; t.explanation = 'Hpが全快した'; break;
        case 1: p.atk = cLng(p.atk * 1.1 + 1); t.explanation = '攻撃力が上がった'; break;
        case 2: killAll(); t.explanation = 'モンスターが全滅した。'; break;
        case 3:
          for (const m of ms) m.ability = SLOW;
          t.explanation = 'モンスターの動きが遅くなった。';
          break;
        case 4: p.maxHp = cLng(p.maxHp * 1.1 + 1); t.explanation = '最大Hpが上がった'; break;
        case 5:
          for (let i = 1; i <= LAND_NUMBER - 2; i++) {
            for (let j = 1; j <= LAND_NUMBER - 2; j++) {
              if (g.land[i][j].condition === WALL) g.land[i][j].condition = ROOM;
            }
          }
          t.explanation = '一番外側以外の壁が全て崩れた。';
          break;
        case 6: p.condition = STEALTH; t.explanation = 'プレイヤーは透明になった。(このフロアのみ)'; break;
        case 7:
          for (const m of ms) m.hp = 1;
          t.explanation = '全てのモンスターのHpが残り1になった';
          break;
        case 8:
          for (const m of ms) m.atk = cLng(m.atk * 0.9);
          t.explanation = '全てのモンスターの攻撃力が少し下がった';
          break;
        case 9:
          for (const m of ms) m.exp += 5;
          t.explanation = '全てのモンスターの経験値が少し上がった';
          break;
        case 10:
          for (const m of ms) m.maxHp = cLng(m.maxHp * 0.9);
          t.explanation = '全てのモンスターの最大Hpが少し下がった';
          break;
        case 11:
          eachBox(isBox, (b) => { b.hp = 1; });
          t.explanation = '全ての箱のHpが残り1になった。';
          break;
        case 12:
          eachBox((c) => c === RED_BOX, (b) => { b.condition = ROOM; });
          t.explanation = '赤箱が消滅した。';
          break;
        case 13: p.def++; t.explanation = '守備力が少し上がった'; break;
        case 14: p.condition = WALL_BREAK; t.explanation = 'このフロアにいる間、壁を掘れるようになった(一番外側を除く)'; break;
      }
      break;

    case PURPLE_BOX: {
      const n = vbInt(rnd() * (5 - p.ability)) + 1;
      const a = t.ability;
      if (a <= 2) { g.abilityHp[0] += n; t.explanation = `全体攻撃の使用回数が${n}増えた。`; }
      else if (a <= 6) { g.abilityHp[1] += n; t.explanation = `Hp全快の使用回数が${n}増えた。`; }
      else if (a === 7) { g.abilityHp[2] += n; t.explanation = `全消去の使用回数が${n}増えた。`; }
      else if (a === 8) { g.abilityHp[3] += n; t.explanation = `モンスター箱化の使用回数が${n}増えた。`; }
      else if (a <= 11) { g.abilityHp[4] += n; t.explanation = `次の階へ行く効果の使用回数が${n}増えた。`; }
      else { g.abilityHp[5] += n; t.explanation = `復活の珠が${n}個手に入った。`; }
      break;
    }

    case WALL:
      t.explanation = '';
      break;
  }

  if (box <= PURPLE_BOX) g.stats.boxes[box]++;
  else g.stats.walls++;
  emit('box', { x, y, box, text: t.explanation });

  statusCheck(false);
  g.words[1] = t.explanation;
  t.ability = 0;
  t.condition = ROOM;
}

// ---------------------------------------------------------------- AbilityModule

function abilityEffect(key) {
  const p = g.player;
  const ah = g.abilityHp;
  switch (key) {
    case 'A': // 隠しコマンド: 緑箱を全部開けて次の階へ
      for (let x = 0; x < LAND_NUMBER; x++) {
        for (let y = 0; y < LAND_NUMBER; y++) {
          if (g.land[x][y].condition === GREEN_BOX) {
            g.land[x][y].ability = vbInt(rnd() * 15);
            g.turn += 10;
            boxEffect(x, y);
          }
        }
      }
      if (g.mode === DUNGEON) floorSet();
      break;

    case 'Z': // 全体攻撃
      if (ah[0] > 0) {
        ah[0]--;
        g.stats.commands++;
        emit('cmd', { key });
        let damage = vbInt((p.atk * (rnd() * 0.2 + 0.9)) / g.monsters[0].def);
        if (damage > LIMIT_7) damage = LIMIT_7;
        for (let i = 0; i < g.monsterNumber; i++) {
          const m = g.monsters[i];
          if (!m.alive) continue;
          m.hp -= damage;
          emit('hit', { x: m.x, y: m.y, dmg: damage });
          if (m.hp <= 0) {
            m.hp = 0;
            m.alive = false;
            g.land[m.x][m.y].condition = ROOM;
            p.exp += m.exp;
            g.stats.kills++;
            emit('kill', { x: m.x, y: m.y });
          }
          g.words[2] = `全てのモンスターは${damage}ダメージを受けた`;
        }
        p.direction = DIR_IDLE; // 1 ターン消費
      }
      break;

    case 'X': // Hp全快
      if (ah[1] > 0) {
        ah[1]--;
        p.hp = p.maxHp;
        g.stats.commands++;
        emit('cmd', { key });
      }
      break;

    case 'C': // 全消去
      if (ah[2] > 0) {
        ah[2]--;
        g.stats.commands++;
        emit('cmd', { key });
        for (const col of g.land) {
          for (const t of col) {
            if (t.condition <= GREEN_BOX || t.condition === WALL) t.condition = ROOM;
          }
        }
        for (let i = 0; i < g.monsterNumber; i++) {
          const m = g.monsters[i];
          if (m.alive) {
            m.alive = false;
            m.hp = 0;
            m.condition = 0;
            g.land[m.x][m.y].condition = ROOM;
          }
        }
      }
      break;

    case 'D':
      if (ah[3] > 0 && p.alive) { // モンスター箱化
        ah[3]--;
        g.stats.commands++;
        emit('cmd', { key });
        for (let i = 0; i < g.monsterNumber; i++) {
          const m = g.monsters[i];
          if (m.alive) {
            m.alive = false;
            m.hp = 0;
            m.condition = 0;
            g.land[m.x][m.y].condition = vbInt(rnd() * 4);
          }
        }
      } else if (ah[5] > 0 && !p.alive) { // 復活の珠
        ah[5]--;
        p.alive = true;
        p.hp = p.maxHp;
        if (p.condition !== TURN_CONST) g.turn += 100;
        p.condition = NO_ABILITY;
        g.words[3] = '復活の珠の効果で復活した。';
        g.stats.revives++;
        emit('revive');
        statusCheck(false);
      }
      break;

    case 'S': // セーブ予約 / 階段でセーブ
      if (ah[6] === 0) {
        ah[6]++;
        g.words[3] = '次に階段を降りた時、セーブしてメニュー画面に戻ります。';
      } else if (g.land[p.x][p.y].condition === STAIR) {
        statusCheck(false);
        const m0 = g.monsters[0];
        storage.save({
          player: {
            hp: p.hp, maxHp: p.maxHp, atk: p.atk, def: p.def,
            exp: p.exp, level: p.level, ability: p.ability,
          },
          abilityHp: ah.slice(0, 6),
          floor: g.floor,
          turn: g.turn,
          monster: {
            maxHp: m0.maxHp, atk: m0.atk, def: m0.def,
            exp: m0.exp, level: m0.level, ability: m0.ability,
          },
          monsterNumber: g.monsterNumber,
          arrange: g.arrange,
          stats: g.stats,
        });
        g.mode = ENTRANCE;
        landSet();
        wordSet();
        g.music = -1; // メニュー曲を頭から
      }
      break;

    case 'L': { // ロード（Museum から）
      const d = storage.load();
      if (!d) break;
      g.mode = DUNGEON;
      landSet();
      monsterSet();
      // セーブした冒険の続きとして集計とルールを引き継ぐ
      g.stats = { ...makeStats(), ...d.stats, boxes: [...(d.stats?.boxes ?? [0, 0, 0, 0, 0])] };
      g.result = null;
      if (typeof d.arrange === 'boolean') g.arrange = d.arrange;
      for (const m of g.monsters) {
        m.maxHp = d.monster.maxHp;
        m.hp = m.maxHp;
        m.atk = d.monster.atk;
        m.def = d.monster.def;
        m.exp = d.monster.exp;
        m.level = d.monster.level;
        m.ability = d.monster.ability;
      }
      floorSet();
      break;
    }

    case 'Enter': // 次の階へ
      if (ah[4] > 0) {
        ah[4]--;
        g.stats.commands++;
        emit('cmd', { key });
        floorSet();
      }
      break;
  }
}

// ---------------------------------------------------------------- WordModule

// GameOver / GameClear 画面の冒険の結果（w[from]..w[from+3]）
function resultWords(from) {
  const w = g.words;
  const r = g.result;
  const s = g.stats;
  for (let i = from; i < from + 4; i++) w[i] = '';
  if (!r) return;
  const b = s.boxes;
  w[from] = `到達 ${r.floor}F  Lv${r.level}${r.auto ? '  (オートプレイ)' : ''}\n称号「${titleFor(r.floor)}」`;
  w[from + 1] = r.newBest ? '★ 自己ベスト更新！ ★' : `自己ベスト ${g.records.best}F`;
  w[from + 2] = `壊した箱 ${b[0] + b[1] + b[2] + b[3] + b[4]} (青${b[0]} 赤${b[1]} 黄${b[2]} 緑${b[3]} 紫${b[4]})`;
  w[from + 3] = `倒したモンスター ${s.kills}  最大ダメージ ${s.maxDamage}`;
}

// 記録の部屋の冒険の記録（w[5]..w[8]）
function recordWords() {
  const w = g.words;
  const r = g.records;
  w[5] = '';
  w[6] = r.best > 0 ? `最高記録 ${r.best}F「${titleFor(r.best)}」` : '冒険の記録はまだありません。';
  w[7] = r.runs.slice(0, 3).map((run, i) =>
    `${i + 1}. ${run.floor}F Lv${run.level} ${run.arrange ? 'アレンジ' : 'オリジナル'}${run.auto ? ' (オート)' : ''}`).join('\n');
  w[8] = '';
}

const BACK_HELP = '(xキーでメニュー画面に戻る,左右キーでページ選択)';

const HOWTO_PAGES = [
  [
    'How to play\n1/7',
    '\nこのゲームは勘でも何とかなるゲームです。',
    'なので、説明を読むのが嫌いな人は、\nここは読み飛ばしても大丈夫です。',
    '\n', '\n', '\n', '\n', '\n',
    `\n${BACK_HELP}`,
  ],
  [
    'How to play\n2/7',
    'このゲームはモンスターを倒しながら、\n下の階を目指して階段を降りていくゲームです。',
    'プレイヤー　　　                     \nモンスター　　　                     ',
    '階段　　　                     \n(階段はマウスポインタと見間違えやすいので注意して下さい。)',
    '操作は基本的に十字キーによる移動だけです。\n攻撃も相手の居るマスに進もうとするだけで出来ます。',
    '\n壁　　　                     　',
    '床　　　                     　\n',
    'ダンジョンの地形は上の2種類だけです。\nこのうち壁の上には進むことが出来ません。',
    `\n${BACK_HELP}`,
  ],
  [
    'How to play\n3/7',
    'ダンジョンに出現する5つの箱\n',
    '青箱・・・壊すとHPが回復し、モンスターが普通の状態に戻る。\n赤箱・・・壊すとマイナス効果が発生する。',
    '黄箱・・・壊すと色々な効果が発生する\n緑箱・・・壊すとプラス効果が発生する。',
    '紫箱・・・壊すとコマンドキーの使用回数が増える。\n',
    '紫箱について\n',
    '上の5つの箱のうち、紫箱は少し特殊です。\nこの箱は効果などで別の箱に変わらず、',
    '難易度によって効果が違います。\n',
    `\n${BACK_HELP}`,
  ],
  [
    'How to play\n4/7',
    'ダンジョン内で使える5つのコマンド\n',
    'zキー...敵全員に攻撃\nxキー...Hpを全快させる',
    'cキー...紫箱以外の全ての箱、壁、モンスターを消去する。\ndキー...モンスターを箱に変化させる',
    'Enterキー...一つ下のフロアに降りる\n',
    '(これらのことはゲーム中表示されているので、\nどのキーがどんな効果かは覚えなくて大丈夫です。)',
    '\n', '\n',
    `\n${BACK_HELP}`,
  ],
  [
    'How to play\n5/7',
    'セーブ/ロードについて\n',
    'ゲームを中断したくなった時はSボタンを押すとセーブが出来ます。\n再開したい時は記録の部屋からロードできます。',
    'セーブすると、以前のデータは消えてしまうので、注意して下さい。\n',
    'ターン数について\n',
    'このゲームでは1ターンごとにターン数が1減っていきます。\nターン数が0になるとゲームオーバーです。',
    '次のことをすると、ターン数は回復します。\n',
    '箱を壊すか、壁を掘る(能力を手に入れた時のみ)。\n復活の珠で復活する。',
    `\n${BACK_HELP}`,
  ],
  [
    'How to play\n6/7',
    'ダンジョンで大事な6つのステータス\n',
    'レベル\nこれが上がると、全体的に強くなります。',
    '攻撃力\nこの数値が大きいほど、与えるダメージが大きくなります。',
    '守備力\nこの数値が大きいほど、受けるダメージが少なくなります。',
    'Hp\n体力です。これが0になると倒されます。',
    '最大Hp\nHpの最大値です。Hpはこれ以上には回復しません。',
    '経験値\nこれが貯まるとレベルアップしていきます。',
    `モンスターを倒すと、その経験値が自分のものになります。\n${BACK_HELP}`,
  ],
  [
    'How to play\n7/7',
    'ブラウザ版のアレンジ要素\n',
    'ダンジョン・・・フロアの形と箱の並びが毎回変わります。\n序盤は青箱や緑箱が多く、深い階ほど赤箱が増えます。',
    '祝福・・・10階を越えるごとに、3つの祝福から1つを選べます。\n(左右キーで選択、Enterキーで決定)',
    'Optionでルールを「オリジナル」にすると、元のダンジョンになり、\n祝福も現れません。',
    '記録と称号\n',
    '冒険が終わると、到達した階が記録の部屋に残ります。\n最も深く潜った階に応じて称号が付きます。',
    '画面右下には攻撃力・守備力と、\nモンスターの強さが表示されています。',
    `\n${BACK_HELP}`,
  ],
];

export function wordSet() {
  const w = g.words;
  const p = g.player;
  const ah = g.abilityHp;
  switch (g.mode) {
    case ENTRANCE:
      w[0] = 'ダンジョンと不思議の箱\n';
      w[1] = 'ダンジョンに入る(Enterキー)\n不思議の箱を駆使する冒険が始まります。\n';
      w[2] = 'How to play(zキー)\n操作説明や概要説明など\n';
      w[3] = 'Option(cキー)\n難易度を設定\n';
      w[4] = `Museum(dキー)\n記録の部屋\n${g.records.best > 0 ? `最高記録 ${g.records.best}F「${titleFor(g.records.best)}」` : ''}\n(Mキーで音楽のON/OFF)`;
      w[5] = '\n音楽素材提供元';
      w[6] = '【サイト名】フリー音楽素材 H/MIX GALLERY\n【管理者】　秋山裕和';
      w[7] = '【アドレス】http://www.hmix.net/\n';
      w[8] = '\n';
      break;

    case DUNGEON:
      w[4] = `全体攻撃(z)\n${ah[0]}`;
      w[5] = `Hp全快(x)\n${ah[1]}`;
      w[6] = `全消去(c)\n${ah[2]}`;
      w[7] = `モンスター箱化(d)\n${ah[3]}`;
      w[8] = `次の階へ(Enter)\n${ah[4]}\n復活の珠\n${ah[5]}`;
      w[9] = `${g.floor}F Lv${p.level} HP${vbInt(p.hp)}/${p.maxHp}  Turn ${g.turn}`;
      // 上 4 行のメッセージは 40 フレーム（2 秒）で消える
      for (let i = 0; i < g.owords.length; i++) {
        const o = g.owords[i];
        if (o.explanation !== w[i]) {
          o.explanation = w[i];
          o.hp = 40;
        } else if (o.hp > 0) {
          o.hp--;
        } else {
          w[i] = '';
          // port: VB6 では同じ文が続くと 2 回目以降が表示されなかったので、消した時に記憶もリセット
          o.explanation = '';
        }
      }
      break;

    case HOWTO_PLAY: {
      const page = HOWTO_PAGES[p.condition] || HOWTO_PAGES[0];
      for (let i = 0; i < 9; i++) w[i] = page[i];
      break;
    }

    case OPTIONS:
      w[0] = 'Options\n';
      w[1] = `${DIFFICULTY_NAMES[p.ability]}\n`;
      w[2] = '低い難易度ほど紫箱で良い効果が出やすいです。\n';
      w[3] = `ルール: ${g.arrange ? 'アレンジ' : 'オリジナル'}\n${g.arrange ? 'ダンジョンが毎回変わり、10階ごとに祝福を選べます。' : 'VB6版そのままのルールです。'}\n`;
      w[4] = 'xキーでメニュー画面に戻る\n左右キーで難易度選択,上下キーでルール選択\nEnterキーでダンジョンに入る。';
      for (let i = 5; i <= 8; i++) w[i] = '';
      break;

    case GAME_OVER:
      g.foreColor = '#ffffff';
      w[0] = 'GameOver';
      w[1] = '';
      w[2] = g.result?.cause === 'turn' ? 'ターン数が尽きた' : 'プレイヤーは力尽きた';
      w[3] = '';
      resultWords(4);
      w[8] = '\nxキーでメニュー画面に戻る';
      break;

    case GAME_CLEAR:
      g.backColor = '#ffffff';
      g.foreColor = '#0000ff';
      w[0] = 'GameClear\nここが最下層の1000Fです。';
      w[1] = '\nこのゲームをここまで遊んでくれたあなたは果報者です。';
      w[2] = 'クリア時のステータス\n';
      w[3] = `HP ${vbInt(p.hp)}/${p.maxHp}\nLv ${p.level}`;
      resultWords(4);
      w[8] = '\nxキーでメニュー画面に戻る';
      break;

    case MUSEUM:
      w[0] = '現在の記録\n';
      recordWords();
      if (!g.saveExists) {
        w[1] = '\nセーブデータがありません。';
        w[2] = 'ダンジョン内でSキーを押してから階段を降りると\nセーブできます。';
        w[3] = '';
        w[4] = '\nxキーでメニュー画面に戻る';
        break;
      }
      w[1] = `${g.floor + 1}F Lv${p.level}\nHP${vbInt(p.hp)}/${p.maxHp}  Turn ${g.turn}`;
      w[2] = `全体攻撃 ${ah[0]}\nHp全快 ${ah[1]}\n全消去 ${ah[2]}`;
      w[3] = `モンスター箱化 ${ah[3]}\n次の階へ ${ah[4]}\n復活の珠 ${ah[5]}`;
      w[4] = 'Lキーを押すと、このセーブデータから始められます。\nxキーでメニュー画面に戻る';
      break;
  }
}

