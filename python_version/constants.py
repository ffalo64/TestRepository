"""
Game Constants and Enumerations
ゲームの定数と列挙型の定義
"""
from enum import IntEnum


# Game Modes
class GameMode(IntEnum):
    """ゲームモード"""
    ENTRANCE = 0      # タイトル画面
    DUNGEON = 1       # ダンジョン（メインゲーム）
    HOW_TO_PLAY = 2   # 遊び方
    OPTIONS = 3       # オプション
    GAME_OVER = 4     # ゲームオーバー
    GAME_CLEAR = 5    # ゲームクリア
    MUSEUM = 6        # 記録の部屋


# Land/Tile Types
class TileType(IntEnum):
    """タイルの種類"""
    BLUE_BOX = 0      # 青箱（HP小回復）
    RED_BOX = 1       # 赤箱（マイナス効果）
    YELLOW_BOX = 2    # 黄箱（ランダム効果）
    GREEN_BOX = 3     # 緑箱（プラス効果）
    PURPLE_BOX = 4    # 紫箱（アビリティチャージ）
    STAIR = 5         # 階段
    WALL = 6          # 壁
    ROOM = 7          # 部屋（通路）
    ENEMY = 8         # 敵
    MINE = 9          # 地雷


# Monster Abilities
class MonsterAbility(IntEnum):
    """モンスターの特殊能力"""
    NO_ABILITY = 0    # 能力なし
    WALL_BREAK = 1    # 壁破壊
    SLOW = 2          # スロー（移動速度半分）
    BOX_ATTACK = 3    # 箱攻撃
    STEALTH = 4       # ステルス（透明）


# Player Conditions
class PlayerCondition(IntEnum):
    """プレイヤーの状態"""
    NORMAL = 0        # 通常
    WALL_BREAK = 1    # 壁破壊可能
    SLOW = 2          # スロー状態
    TURN_CONST = 3    # ターン減少なし
    STEALTH = 4       # ステルス状態


# Difficulty Levels
class Difficulty(IntEnum):
    """難易度"""
    VERY_EASY = 0
    EASY = 1
    NORMAL = 2
    HARD = 3
    VERY_HARD = 4


# Direction multipliers (prime number system from VB6)
class Direction:
    """方向の定数（素数の積で方向を管理）"""
    NONE = 1
    UP = 2
    DOWN = 3
    RIGHT = 5
    LEFT = 7


# Map constants
LAND_NUMBER = 40      # マップのサイズ（40x40）
SC = 10               # 1マスのピクセルサイズ
PLAYER_CONST = 10**9  # プレイヤー識別用の定数

# Game balance constants
MAX_MONSTERS = 1000    # 最大モンスター数
INITIAL_TURN = 200     # 初期ターン数
GOAL_FLOOR = 1000      # ゴールの階数

# Status limits
MAX_ATK = 10**9 - 1
MAX_DEF = 10**9 - 1
MAX_EXP = 10**9 - 1
MAX_PLAYER_HP = 10**7 - 1
MAX_MONSTER_HP = 10**6 - 1
MAX_TURN = 10**4 - 1
MAX_DAMAGE = 10**7 - 1

# Color constants (RGB)
class Color:
    """色定数"""
    BLACK = (0, 0, 0)
    WHITE = (255, 255, 255)
    RED = (255, 0, 0)
    BLUE = (0, 0, 255)
    GREEN = (34, 177, 76)
    ORANGE = (255, 127, 39)
    PURPLE = (128, 0, 128)
    YELLOW = (255, 255, 0)

# Sound file paths (relative to project root)
SOUND_MENU = "Sounds/n31.mp3"       # メニュー画面
SOUND_CLEAR = "Sounds/n11.mp3"      # ゲームクリア
SOUND_OVER = "Sounds/c26.mp3"       # ゲームオーバー
SOUND_DUNGEON = "Sounds/c1.mp3"     # ダンジョン

# Image file paths (for future use if we want to load actual sprites)
IMAGE_MAP = "Images/Map40.png"
IMAGE_PLAYER = "Images/Player.bmp"
IMAGE_MONSTER = "Images/Monster.bmp"
IMAGE_BOX = "Images/Box.bmp"
IMAGE_WALL = "Images/Wall.bmp"
IMAGE_ROOM = "Images/Room.bmp"
IMAGE_HPBAR = "Images/HpBar.bmp"
