export const SC = 10;
export const LAND_NUMBER = 40;
export const PLAYER_CONST = 1e9;

export const GameMode = Object.freeze({ Entrance:0, Dungeon:1, HowtoPlay:2, Options:3, GameOver:4, GameClear:5, Museum:6 });

// Terrain / LandSquare conditions
export const T = Object.freeze({ BlueBox:0, RedBox:1, YellowBox:2, GreenBox:3, PurpleBox:4, Stair:5, Wall:6, Room:7, Enemy:8, Mine:9 });

// Monster / player ability (condition)
export const A = Object.freeze({ Noability:0, WallBreak:1, Slow:2, Boxattack:3, Stealth:4 });

// TurnConst = 3 (same value as Boxattack but used for player.condition)
export const TURN_CONST = 3;

export const CANVAS_W = 600;
export const CANVAS_H = 600;
export const MAP_W = LAND_NUMBER * SC; // 400

// Layout constants
export const SPEC_X = MAP_W;            // 400: SpecBar x start
export const STATUS_Y = MAP_W;          // 400: StatusBar y start
export const MSG_Y = STATUS_Y + 50;     // 450: MessageBar y start
export const TILE_PX = SC;             // render tile size = SC (no scaling, map fits 400×400)
