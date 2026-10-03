// VB6 の Public Const を集約（BasicModule / LandModule / MonsterModule / AbilityModule）

// GameMode
export const ENTRANCE = 0;
export const DUNGEON = 1;
export const HOWTO_PLAY = 2;
export const OPTIONS = 3;
export const GAME_OVER = 4;
export const GAME_CLEAR = 5;
export const MUSEUM = 6;
export const BLESSING = 7; // ブラウザ版独自: 祝福の選択画面

// Landsquare.Condition
export const BLUE_BOX = 0;
export const RED_BOX = 1;
export const YELLOW_BOX = 2;
export const GREEN_BOX = 3;
export const PURPLE_BOX = 4;
export const STAIR = 5;
export const WALL = 6;
export const ROOM = 7;
export const ENEMY = 8;
export const MINE = 9;

export const LAND_NUMBER = 40;
export const SC = 10; // Square Const（1マスのピクセル数）

// Monster.Ability / Player.Condition
export const NO_ABILITY = 0;
export const WALL_BREAK = 1;
export const SLOW = 2;
export const BOX_ATTACK = 3;
export const TURN_CONST = 3;
export const STEALTH = 4;

export const MONSTER_MAX = 1000; // ReDim Monster(999)

// 難易度（Player.Ability）
export const DIFFICULTY_NAMES = ['Very Easy', 'Easy', 'Normal', 'Hard', 'Very Hard'];

// フォーム: 600x600, メイリオ 14.25pt 太字
export const SCREEN = 600;
export const FONT_SIZE = 14.25;
export const LINE_HEIGHT = FONT_SIZE * 2;

// 方向は素数の積（1=なし, 2=上, 3=下, 5=右, 7=左, 11=鈍足の空ターン）
export const DIR_NONE = 1;
export const DIR_UP = 2;
export const DIR_DOWN = 3;
export const DIR_RIGHT = 5;
export const DIR_LEFT = 7;
export const DIR_IDLE = 11;

// 音楽トラック番号（Lib.SoundSet の番号）
export const MUSIC_MENU = 0;
export const MUSIC_CLEAR = 1;
export const MUSIC_GAMEOVER = 2;
export const MUSIC_DUNGEON = 3;
