"""
Game Data Models
ゲームのデータモデル定義
"""
from dataclasses import dataclass, field
from typing import Optional
from constants import PlayerCondition, MonsterAbility, TileType


@dataclass
class Status:
    """
    ゲーム内のオブジェクトのステータス
    プレイヤー、モンスター、タイルなどで使用
    """
    # Position
    left: int = 0
    top: int = 0
    oleft: int = 0  # Old left (previous position)
    otop: int = 0   # Old top (previous position)

    # Size
    width: int = 10
    height: int = 10

    # Status
    alive: bool = False
    hp: int = 0
    max_hp: int = 0
    exp: int = 0
    atk: int = 0
    defense: int = 0
    level: int = 1

    # State
    condition: int = 0
    direction: int = 1
    ability: int = 0

    # Text
    name: str = ""
    explanation: str = ""


@dataclass
class PlayerStatus:
    """プレイヤー専用のステータス"""
    # Position
    left: int = 0
    top: int = 0
    oleft: int = 0
    otop: int = 0

    # Size
    width: int = 10
    height: int = 10

    # Core stats
    alive: bool = True
    hp: int = 20
    max_hp: int = 20
    exp: int = 0
    atk: int = 10
    defense: int = 3
    level: int = 1

    # State
    condition: PlayerCondition = PlayerCondition.NORMAL
    direction: int = 1
    ability: int = 0

    def reset(self):
        """初期状態にリセット"""
        self.hp = 20
        self.max_hp = 20
        self.level = 1
        self.exp = 0
        self.atk = 10
        self.defense = 3
        self.direction = 1
        self.condition = PlayerCondition.NORMAL
        self.alive = True


@dataclass
class MonsterStatus:
    """モンスター専用のステータス"""
    # Position
    left: int = 0
    top: int = 0
    oleft: int = 0
    otop: int = 0

    # Size
    width: int = 10
    height: int = 10

    # Core stats
    alive: bool = False
    hp: int = 10
    max_hp: int = 10
    exp: int = 10
    atk: int = 10
    defense: int = 1
    level: int = 1

    # State
    condition: int = 0
    direction: int = 1
    ability: MonsterAbility = MonsterAbility.NO_ABILITY

    # Text
    name: str = ""
    explanation: str = ""


@dataclass
class LandSquare:
    """マップの1マス分のデータ"""
    # Position
    left: int = 0
    top: int = 0

    # Size
    width: int = 10
    height: int = 10

    # Status
    alive: bool = True
    hp: int = 0
    max_hp: int = 0
    defense: int = 1

    # State
    condition: TileType = TileType.ROOM
    ability: int = 0
    direction: int = 0  # Used for color storage in MapSet

    # Text
    name: str = ""
    explanation: str = ""


@dataclass
class GameState:
    """ゲーム全体の状態管理"""
    game_mode: int = 0  # GameMode
    floor: int = 0
    turn: int = 200
    sum_damage: int = 0
    monster_number: int = 10
    difficulty: int = 2  # Normal

    # Ability usage counts
    ability_hp: list[int] = field(default_factory=lambda: [0] * 7)

    # Messages
    words: list[str] = field(default_factory=lambda: [""] * 12)

    def reset(self):
        """ゲーム状態をリセット"""
        self.floor = 0
        self.turn = 200
        self.sum_damage = 0
        self.monster_number = 10
        self.ability_hp = [0] * 7
        self.words = [""] * 12


@dataclass
class Rectangle:
    """矩形領域の定義（UIバー用）"""
    left: int = 0
    top: int = 0
    right: int = 0
    bottom: int = 0

    @property
    def width(self) -> int:
        return self.right - self.left

    @property
    def height(self) -> int:
        return self.bottom - self.top
