#!/usr/bin/env python3
"""
ダンジョンと不思議の箱 - Python版ゲームエンジン

Visual Basicで開発されたオリジナルゲーム「ダンジョンと不思議の箱」を
Pythonでリファクタリングした実装です。
"""

from dataclasses import dataclass, field
from enum import Enum, IntEnum
from typing import List, Tuple, Optional
import random


# ============================================================================
# 列挙型定義 (Enums)
# ============================================================================

class GameMode(IntEnum):
    """ゲームモード"""
    ENTRANCE = 0      # エントランス
    DUNGEON = 1       # ダンジョン
    HOW_TO_PLAY = 2   # 操作説明
    OPTIONS = 3       # オプション
    GAME_OVER = 4     # ゲームオーバー
    GAME_CLEAR = 5    # ゲームクリア
    MUSEUM = 6        # 記録閲覧


class LandCondition(IntEnum):
    """地形の状態"""
    BLUE_BOX = 0      # 青箱
    RED_BOX = 1       # 赤箱
    YELLOW_BOX = 2    # 黄箱
    GREEN_BOX = 3     # 緑箱
    PURPLE_BOX = 4    # 紫箱
    STAIR = 5         # 階段
    WALL = 6          # 壁
    ROOM = 7          # 部屋
    ENEMY = 8         # 敵
    MINE = 9          # 地雷


class Ability(IntEnum):
    """特殊能力"""
    NO_ABILITY = 0    # 能力なし
    WALL_BREAK = 1    # 壁破壊
    SLOW = 2          # 移動速度低下
    BOX_ATTACK = 3    # 箱攻撃
    TURN_CONST = 3    # ターン消費なし
    STEALTH = 4       # 透明


class Difficulty(IntEnum):
    """難易度"""
    VERY_EASY = 0
    EASY = 1
    NORMAL = 2
    HARD = 3
    VERY_HARD = 4


# ============================================================================
# 基本データ構造 (Data Classes)
# ============================================================================

@dataclass
class Position:
    """位置情報"""
    x: int = 0
    y: int = 0


@dataclass
class Rectangle:
    """矩形領域"""
    left: int = 0
    top: int = 0
    right: int = 0
    bottom: int = 0


@dataclass
class Status:
    """基本ステータス"""
    left: int = 0
    top: int = 0
    oleft: int = 0          # 前回のLeft位置
    otop: int = 0           # 前回のTop位置
    width: int = 10
    height: int = 10
    alive: bool = False
    hp: int = 0
    max_hp: int = 0
    exp: int = 0            # 経験値
    atk: int = 0            # 攻撃力
    ability: int = 0        # 特殊能力
    defense: int = 0        # 防御力
    level: int = 1
    condition: int = 0      # 状態
    direction: int = 1      # 方向
    name: str = ""
    explanation: str = ""   # 説明文


@dataclass
class Player(Status):
    """プレイヤークラス"""
    pass


@dataclass
class Monster(Status):
    """モンスタークラス"""
    pass


@dataclass
class LandSquare(Status):
    """地形スクエアクラス"""
    pass


# ============================================================================
# ゲーム状態管理 (Game State)
# ============================================================================

class GameState:
    """ゲーム状態を管理するクラス"""

    # ゲーム定数
    LAND_NUMBER = 40        # マップサイズ (40x40)
    SC = 10                 # スクエアサイズ
    PLAYER_CONST = 10 ** 9  # プレイヤー識別用定数

    def __init__(self):
        """初期化"""
        # ゲーム状態
        self.game_mode: GameMode = GameMode.ENTRANCE
        self.floor: int = 0
        self.turn: int = 200
        self.sum_damage: int = 0
        self.monster_number: int = 10

        # プレイヤー
        self.player: Player = Player()

        # モンスター配列（最大1000体）
        self.monsters: List[Monster] = [Monster() for _ in range(1000)]

        # 地形マップ (40x40)
        self.land_squares: List[List[LandSquare]] = [
            [LandSquare() for _ in range(self.LAND_NUMBER)]
            for _ in range(self.LAND_NUMBER)
        ]

        # ランダム配置用の一時ステータス
        self.random_position: Status = Status()

        # 特殊能力の使用回数
        self.ability_hp: List[int] = [0] * 7

        # メッセージ配列
        self.words: List[str] = [""] * 12
        self.oword: List[Status] = [Status() for _ in range(4)]

        # 難易度
        self.difficulty: Difficulty = Difficulty.NORMAL

        # UI用の矩形
        self.spec_bar: Rectangle = Rectangle()
        self.status_bar: Rectangle = Rectangle()
        self.message_bar: Rectangle = Rectangle()

    def reset_entrance(self):
        """エントランス画面の初期化"""
        self.game_mode = GameMode.ENTRANCE

        # 能力使用回数の初期化
        self.ability_hp = [0] * 7

        # プレイヤー初期化
        self.player = Player(
            alive=False,
            height=self.SC,
            width=self.SC,
            level=1,
            max_hp=20,
            hp=20,
            condition=0,
            direction=1
        )

        # 地形マップ初期化
        for i in range(self.LAND_NUMBER):
            for j in range(self.LAND_NUMBER):
                self.land_squares[i][j] = LandSquare(
                    width=self.SC,
                    height=self.SC,
                    left=self.SC * i,
                    top=self.SC * j,
                    condition=LandCondition.ROOM,
                    alive=True
                )

        self.floor = 0

    def reset_dungeon(self):
        """ダンジョンの初期化"""
        self.game_mode = GameMode.DUNGEON

        # メッセージクリア
        self.words = [""] * 12

        # プレイヤー復活
        self.player.hp = self.player.max_hp
        self.player.alive = True

    def initialize_monsters(self):
        """モンスターの初期化"""
        for i in range(1000):
            self.monsters[i] = Monster(
                name=f"ヘキサスライム{i}",
                explanation="",
                hp=10,
                max_hp=10,
                level=1,
                exp=10,
                ability=Ability.NO_ABILITY,
                atk=10,
                defense=1,
                height=self.SC,
                width=self.SC,
                direction=1
            )


# ============================================================================
# ゲームエンジン (Game Engine)
# ============================================================================

class GameEngine:
    """ゲームエンジンクラス"""

    def __init__(self):
        """初期化"""
        self.state = GameState()
        self.state.reset_entrance()
        self.state.initialize_monsters()

    # ========================================================================
    # 移動処理 (Movement)
    # ========================================================================

    def movement(self):
        """移動処理"""
        if self.state.game_mode == GameMode.ENTRANCE:
            self.state.player.direction = 1

        elif self.state.game_mode == GameMode.DUNGEON:
            if self.state.player.direction > 1:
                self._player_movement()
                self._monster_movement()

                if self.state.sum_damage > 0:
                    self.state.words[3] = f"プレイヤーは{self.state.sum_damage}ダメージを受けた"
                    self.state.sum_damage = 0

    def _player_movement(self):
        """プレイヤー移動処理"""
        p = self.state.player
        sc = GameState.SC
        ln = GameState.LAND_NUMBER

        # 方向に応じた移動
        if (p.direction % 2 == 0) and (p.top > 0):  # 上
            p.top -= p.height
        if (p.direction % 3 == 0) and (p.top < ln * sc - p.height):  # 下
            p.top += p.height
        if (p.direction % 5 == 0) and (p.left < ln * sc - p.width):  # 右
            p.left += p.width
        if (p.direction % 7 == 0) and (p.left > 0):  # 左
            p.left -= p.width

        # 位置チェック
        self._position_check(GameState.PLAYER_CONST)

        # ターン処理
        if p.condition != Ability.SLOW:
            p.direction = 1
            self.state.turn -= 1
        elif (p.condition == Ability.SLOW) and (self.state.turn % 2 == 0):
            p.direction = 11
            self.state.turn -= 1
        else:
            p.direction = 1
            self.state.turn -= 1

    def _monster_movement(self):
        """モンスター移動処理"""
        for i in range(self.state.monster_number):
            m = self.state.monsters[i]
            p = self.state.player

            # 生存中かつ移動可能な状態
            should_move = False
            if m.alive and (m.ability != Ability.SLOW) and (p.condition != Ability.STEALTH):
                should_move = True
            elif m.alive and (m.ability == Ability.SLOW) and (self.state.turn % 2 == 0) and (p.condition != Ability.STEALTH):
                should_move = True

            if should_move:
                # プレイヤーに向かって移動
                if m.top > p.top:
                    m.top -= m.height
                    m.direction *= 2
                if m.top < p.top:
                    m.top += m.height
                    m.direction *= 3
                if m.left < p.left:
                    m.left += m.width
                    m.direction *= 5
                if m.left > p.left:
                    m.left -= m.width
                    m.direction *= 7

                self._position_check(i)
                m.direction = 1

    # ========================================================================
    # 位置チェック (Position Check)
    # ========================================================================

    def _position_check(self, character: int):
        """位置チェックと衝突判定"""
        if character == GameState.PLAYER_CONST:
            self._position_check_player()
        else:
            self._position_check_monster(character)

    def _position_check_player(self):
        """プレイヤーの位置チェック"""
        p = self.state.player
        sc = GameState.SC
        ln = GameState.LAND_NUMBER

        x, y = p.left // sc, p.top // sc
        land = self.state.land_squares[x][y]

        # 階段チェック
        if land.condition == LandCondition.STAIR:
            if self.state.ability_hp[6] == 0:
                self._floor_set()
            return

        # 箱・壁チェック
        if land.condition <= LandCondition.PURPLE_BOX:
            self._battle_check(GameState.PLAYER_CONST, x, y)
            p.left, p.top = p.oleft, p.otop
        elif land.condition == LandCondition.WALL:
            if (p.condition == Ability.WALL_BREAK and
                p.left != 0 and p.top != 0 and
                p.left != (ln - 1) * sc and p.top != (ln - 1) * sc):
                self._battle_check(GameState.PLAYER_CONST, x, y)
            p.left, p.top = p.oleft, p.otop

        # モンスターとの衝突チェック
        for i in range(self.state.monster_number):
            m = self.state.monsters[i]
            if p.left == m.left and p.top == m.top and m.alive:
                self._battle_check(GameState.PLAYER_CONST, i, ln)
                p.left, p.top = p.oleft, p.otop

        # 位置更新
        p.oleft, p.otop = p.left, p.top

    def _position_check_monster(self, monster_idx: int):
        """モンスターの位置チェック"""
        m = self.state.monsters[monster_idx]
        if not m.alive:
            return

        p = self.state.player
        sc = GameState.SC
        ln = GameState.LAND_NUMBER

        # プレイヤーとの衝突
        if m.left == p.left and m.top == p.top and p.alive:
            m.left, m.top = m.oleft, m.otop
            self._battle_check(monster_idx, GameState.PLAYER_CONST, ln)
            return

        x, y = m.left // sc, m.top // sc
        land = self.state.land_squares[x][y]

        # 部屋以外の場合
        if land.condition != LandCondition.ROOM:
            # 壁破壊能力
            if (land.condition == LandCondition.WALL and
                m.ability == Ability.WALL_BREAK and
                m.left != 0 and m.top != 0 and
                m.left != (ln - 1) * sc and m.top != (ln - 1) * sc):
                self._battle_check(monster_idx, x, y)
            # 箱攻撃能力
            elif (land.condition <= LandCondition.PURPLE_BOX and
                  m.ability == Ability.BOX_ATTACK):
                self._battle_check(monster_idx, x, y)

            # 壁にぶつかった場合の迂回処理
            self._monster_detour(monster_idx)

        # 位置更新
        old_x, old_y = m.oleft // sc, m.otop // sc
        self.state.land_squares[old_x][old_y].condition = LandCondition.ROOM
        self.state.land_squares[x][y].condition = LandCondition.ENEMY
        m.oleft, m.otop = m.left, m.top

    def _monster_detour(self, monster_idx: int):
        """モンスターの迂回処理"""
        m = self.state.monsters[monster_idx]
        sc = GameState.SC
        old_x, old_y = m.oleft // sc, m.otop // sc

        # 方向に応じて迂回
        detour_map = {
            2: [(old_x - 1, old_y - 1), (old_x + 1, old_y - 1)],
            3: [(old_x + 1, old_y + 1), (old_x - 1, old_y + 1)],
            5: [(old_x + 1, old_y - 1), (old_x + 1, old_y + 1)],
            7: [(old_x - 1, old_y + 1), (old_x - 1, old_y - 1)],
            10: [(old_x, old_y - 1), (old_x + 1, old_y)],
            15: [(old_x + 1, old_y), (old_x, old_y + 1)],
            21: [(old_x, old_y + 1), (old_x - 1, old_y)],
            14: [(old_x - 1, old_y), (old_x, old_y - 1)],
        }

        if m.direction in detour_map:
            for dx, dy in detour_map[m.direction]:
                if 0 <= dx < GameState.LAND_NUMBER and 0 <= dy < GameState.LAND_NUMBER:
                    if self.state.land_squares[dx][dy].condition == LandCondition.ROOM:
                        m.left, m.top = dx * sc, dy * sc
                        return

        # 迂回できない場合は元の位置に戻る
        m.left, m.top = m.oleft, m.otop

    # ========================================================================
    # 戦闘処理 (Battle Check)
    # ========================================================================

    def _battle_check(self, attacker: int, target_i: int, target_j: int = 0):
        """戦闘チェック"""
        if attacker == GameState.PLAYER_CONST:
            self._battle_check_player_attack(target_i, target_j)
        else:
            self._battle_check_monster_attack(attacker, target_i, target_j)

    def _battle_check_player_attack(self, i: int, j: int):
        """プレイヤーの攻撃"""
        p = self.state.player
        ln = GameState.LAND_NUMBER

        # 地形への攻撃
        if j < ln:
            land = self.state.land_squares[i][j]
            damage = int(p.atk * (random.random() * 0.2 + 0.9) / land.defense) + 1
            damage = min(damage, 10 ** 7 - 1)

            land.hp -= damage
            self.state.words[2] = f"{land.name}は{damage}ダメージを受けた"

            if land.hp <= 0:
                land.hp = 0
                land.ability = random.randint(0, 14)
                if p.condition != Ability.TURN_CONST:
                    self.state.turn += 10
                self._box_effect(i, j)

        # モンスターへの攻撃
        else:
            m = self.state.monsters[i]
            if m.alive:
                damage = int(p.atk * (random.random() * 0.2 + 0.9) / m.defense) + 1
                damage = min(damage, 10 ** 7 - 1)

                m.hp -= damage
                self.state.words[2] = f"{m.name}は{damage}ダメージを受けた"

                if m.hp <= 0:
                    m.hp = 0
                    m.alive = False
                    sc = GameState.SC
                    x, y = m.left // sc, m.top // sc
                    self.state.land_squares[x][y].condition = LandCondition.ROOM
                    p.exp += m.exp

    def _battle_check_monster_attack(self, attacker_idx: int, target_i: int, target_j: int):
        """モンスターの攻撃"""
        m = self.state.monsters[attacker_idx]
        p = self.state.player
        ln = GameState.LAND_NUMBER

        # 地形への攻撃
        if target_j < ln:
            land = self.state.land_squares[target_i][target_j]
            if land.condition != LandCondition.STAIR:
                damage = int(m.atk * (random.random() * 0.2 + 0.9) / land.defense) + 1
                damage = min(damage, 10 ** 7 - 1)

                land.hp -= damage
                self.state.words[2] = f"{land.name}は{damage}ダメージを受けた"

                if land.hp <= 0:
                    land.hp = 0
                    land.condition = LandCondition.ROOM

        # プレイヤーへの攻撃
        else:
            damage = int(m.atk * (random.random() * 0.2 + 0.9) / p.defense) + 1
            damage = min(damage, 10 ** 7 - 1)

            p.hp -= damage
            self.state.sum_damage += damage

    # ========================================================================
    # 箱効果 (Box Effect)
    # ========================================================================

    def _box_effect(self, i: int, j: int):
        """箱の効果を適用"""
        land = self.state.land_squares[i][j]
        p = self.state.player

        if land.condition == LandCondition.BLUE_BOX:
            self._box_effect_blue(land)
        elif land.condition == LandCondition.RED_BOX:
            self._box_effect_red(land)
        elif land.condition == LandCondition.YELLOW_BOX:
            self._box_effect_yellow(land)
        elif land.condition == LandCondition.GREEN_BOX:
            self._box_effect_green(land)
        elif land.condition == LandCondition.PURPLE_BOX:
            self._box_effect_purple(land)

        self._status_check(False)
        self.state.words[1] = land.explanation
        land.ability = 0
        land.condition = LandCondition.ROOM

    def _box_effect_blue(self, land: LandSquare):
        """青箱の効果"""
        p = self.state.player
        p.hp += p.max_hp // 10

        for m in self.state.monsters:
            m.ability = Ability.NO_ABILITY

        land.explanation = "HPが少し回復し、全モンスターの状態異常が解除された。"

    def _box_effect_red(self, land: LandSquare):
        """赤箱の効果（15種類のマイナス効果）"""
        p = self.state.player
        effects = [
            lambda: (setattr(p, 'hp', 1), "HPが1になってしまった"),
            lambda: (setattr(p, 'atk', int(p.atk * 0.9)), "攻撃力が下がった"),
            lambda: (setattr(p, 'condition', Ability.SLOW), "プレイヤーはこのフロアにいる間、動きが鈍くなった。"),
            lambda: (setattr(p, 'condition', Ability.TURN_CONST), "このフロアにいる間、ターンが増えなくなった。"),
            lambda: (setattr(p, 'max_hp', int(p.max_hp * 0.9)), "最大HPが下がった"),
            lambda: (setattr(p, 'hp', p.hp // 2 + 1), "HPが半分になった"),
            lambda: self._red_box_effect_6(),
            lambda: self._red_box_effect_7(),
            lambda: self._status_check(True) or "全モンスターのレベルが上がった",
            lambda: self._red_box_effect_9(),
            lambda: self._red_box_effect_10(),
            lambda: self._red_box_effect_11(),
            lambda: self._red_box_effect_12(),
            lambda: self._red_box_effect_13(),
            lambda: self._red_box_effect_14(),
        ]

        result = effects[land.ability]()
        if isinstance(result, tuple):
            land.explanation = result[1]
        elif isinstance(result, str):
            land.explanation = result

    def _red_box_effect_6(self):
        """赤箱効果6: モンスター全回復"""
        for m in self.state.monsters:
            if m.alive:
                m.hp = m.max_hp
                m.condition = 0
        return "全モンスターが全回復した。"

    def _red_box_effect_7(self):
        """赤箱効果7: モンスター攻撃力上昇"""
        for m in self.state.monsters:
            m.atk = int(m.atk * 1.1)
        return "全モンスターの攻撃力が少し上がった"

    def _red_box_effect_9(self):
        """赤箱効果9: 箱の防御力2倍"""
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                if land.condition <= LandCondition.PURPLE_BOX:
                    land.defense *= 2
        return "このフロアの全ての箱の守備力が2倍になった。"

    def _red_box_effect_10(self):
        """赤箱効果10: モンスター防御力上昇"""
        for m in self.state.monsters:
            m.defense = int(m.defense * 1.1)
        return "全モンスターの守備力が少し上がった"

    def _red_box_effect_11(self):
        """赤箱効果11: 全箱が赤箱に"""
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                if land.condition <= LandCondition.GREEN_BOX:
                    land.condition = LandCondition.RED_BOX
        return "全ての箱が赤色になった"

    def _red_box_effect_12(self):
        """赤箱効果12: モンスター最大HP上昇"""
        for m in self.state.monsters:
            m.max_hp = int(m.max_hp * 1.1)
        return "全モンスターの最大HPが少し上がった"

    def _red_box_effect_13(self):
        """赤箱効果13: モンスター経験値減少"""
        for m in self.state.monsters:
            m.exp = int(m.exp * 0.9)
        return "全モンスターの経験値が少し減少した"

    def _red_box_effect_14(self):
        """赤箱効果14: 箱のHP2倍"""
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                if land.condition <= LandCondition.PURPLE_BOX:
                    land.hp *= 2
                    land.max_hp *= 2
        return "このフロアの全ての箱のHPが2倍になった。"

    def _box_effect_yellow(self, land: LandSquare):
        """黄箱の効果（15種類の様々な効果）"""
        effects = [
            lambda: self._yellow_box_effect_0(),
            lambda: self._yellow_box_effect_1(),
            lambda: (setattr(self.state.player, 'condition', Ability.TURN_CONST), "このフロアにいる間、ターンが増えなくなった。"),
            lambda: self._yellow_box_effect_3(),
            lambda: (setattr(self.state.player, 'defense', self.state.player.defense * 2), "守備力が2倍になった"),
            lambda: (setattr(self.state.player, 'atk', self.state.player.atk * 2), "攻撃力が2倍になった"),
            lambda: (setattr(self.state, 'turn', 1000), "ターンの残りが1000になった。"),
            lambda: (setattr(self.state.player, 'atk', self.state.player.atk // 2), "攻撃力が半分になった"),
            lambda: self._yellow_box_effect_8(),
            lambda: self._yellow_box_effect_9(),
            lambda: self._yellow_box_effect_10(),
            lambda: self._yellow_box_effect_11(),
            lambda: self._yellow_box_effect_12(),
            lambda: self._yellow_box_effect_13(),
            lambda: self._yellow_box_effect_14(),
        ]

        result = effects[land.ability]()
        if isinstance(result, tuple):
            land.explanation = result[1]
        elif isinstance(result, str):
            land.explanation = result

    def _yellow_box_effect_0(self):
        """黄箱効果0: モンスター透明化"""
        for m in self.state.monsters:
            m.ability = Ability.STEALTH
        return "全モンスターが透明になった。"

    def _yellow_box_effect_1(self):
        """黄箱効果1: コマンド使用回数+1"""
        for i in range(len(self.state.ability_hp) - 1):
            self.state.ability_hp[i] += 1
        return "全コマンドの使用回数が1増えた。"

    def _yellow_box_effect_3(self):
        """黄箱効果3: モンスター全滅"""
        sc = GameState.SC
        for m in self.state.monsters:
            if m.alive:
                m.alive = False
                m.hp = 0
                m.condition = 0
                x, y = m.left // sc, m.top // sc
                self.state.land_squares[x][y].condition = LandCondition.ROOM
        return "全モンスターが全滅した。"

    def _yellow_box_effect_8(self):
        """黄箱効果8: モンスター壁破壊能力付与"""
        for m in self.state.monsters:
            m.ability = Ability.WALL_BREAK
        return "全モンスターが壁を破壊するようになった。(画面外は例外)"

    def _yellow_box_effect_9(self):
        """黄箱効果9: コマンド使用回数5に"""
        for i in range(len(self.state.ability_hp) - 1):
            self.state.ability_hp[i] = 5
        return "全コマンドの使用回数が5になった。"

    def _yellow_box_effect_10(self):
        """黄箱効果10: 全箱消滅"""
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                if land.condition <= LandCondition.GREEN_BOX:
                    land.condition = LandCondition.ROOM
        return "全ての箱が消えた。"

    def _yellow_box_effect_11(self):
        """黄箱効果11: 全箱が青箱に"""
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                if land.condition <= LandCondition.GREEN_BOX:
                    land.condition = LandCondition.BLUE_BOX
        return "全ての箱が青色になった"

    def _yellow_box_effect_12(self):
        """黄箱効果12: 全箱がランダムに"""
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                if land.condition <= LandCondition.GREEN_BOX:
                    land.condition = random.randint(0, 3)
        return "全ての箱がランダムに変化した。"

    def _yellow_box_effect_13(self):
        """黄箱効果13: 全箱が黄箱に"""
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                if land.condition <= LandCondition.GREEN_BOX:
                    land.condition = LandCondition.YELLOW_BOX
        return "全ての箱が黄色になった"

    def _yellow_box_effect_14(self):
        """黄箱効果14: モンスター箱攻撃能力付与"""
        for m in self.state.monsters:
            m.ability = Ability.BOX_ATTACK
        return "全モンスターが箱を壊せるようになった。"

    def _box_effect_green(self, land: LandSquare):
        """緑箱の効果（15種類のプラス効果）"""
        effects = [
            lambda: (setattr(self.state.player, 'hp', self.state.player.max_hp), "HPが全回復"),
            lambda: (setattr(self.state.player, 'atk', int(self.state.player.atk * 1.1) + 1), "攻撃力が上がった"),
            lambda: self._green_box_effect_2(),
            lambda: self._green_box_effect_3(),
            lambda: (setattr(self.state.player, 'max_hp', int(self.state.player.max_hp * 1.1) + 1), "最大HPが上がった"),
            lambda: self._green_box_effect_5(),
            lambda: (setattr(self.state.player, 'condition', Ability.STEALTH), "プレイヤーは透明になった。(このフロアのみ)"),
            lambda: self._green_box_effect_7(),
            lambda: self._green_box_effect_8(),
            lambda: self._green_box_effect_9(),
            lambda: self._green_box_effect_10(),
            lambda: self._green_box_effect_11(),
            lambda: self._green_box_effect_12(),
            lambda: (setattr(self.state.player, 'defense', self.state.player.defense + 1), "守備力が少し上がった"),
            lambda: (setattr(self.state.player, 'condition', Ability.WALL_BREAK), "このフロアにいる間、壁を破壊するようになった(画面外は例外)"),
        ]

        result = effects[land.ability]()
        if isinstance(result, tuple):
            land.explanation = result[1]
        elif isinstance(result, str):
            land.explanation = result

    def _green_box_effect_2(self):
        """緑箱効果2: モンスター全滅"""
        sc = GameState.SC
        for m in self.state.monsters:
            if m.alive:
                m.alive = False
                m.hp = 0
                m.condition = 0
                x, y = m.left // sc, m.top // sc
                self.state.land_squares[x][y].condition = LandCondition.ROOM
        return "全モンスターが全滅した。"

    def _green_box_effect_3(self):
        """緑箱効果3: モンスター鈍足化"""
        for m in self.state.monsters:
            m.ability = Ability.SLOW
        return "全モンスターの動きが鈍くなった。"

    def _green_box_effect_5(self):
        """緑箱効果5: 壁破壊"""
        for i in range(1, GameState.LAND_NUMBER - 1):
            for j in range(1, GameState.LAND_NUMBER - 1):
                land = self.state.land_squares[i][j]
                if land.condition == LandCondition.WALL:
                    land.condition = LandCondition.ROOM
        return "画面外以外の壁が全て壊れた。"

    def _green_box_effect_7(self):
        """緑箱効果7: モンスターHP1に"""
        for m in self.state.monsters:
            m.hp = 1
        return "全モンスターのHPの残りが1になった"

    def _green_box_effect_8(self):
        """緑箱効果8: モンスター攻撃力減少"""
        for m in self.state.monsters:
            m.atk = int(m.atk * 0.9)
        return "全モンスターの攻撃力が少し減少した"

    def _green_box_effect_9(self):
        """緑箱効果9: モンスター経験値増加"""
        for m in self.state.monsters:
            m.exp += 5
        return "全モンスターの経験値が少し上がった"

    def _green_box_effect_10(self):
        """緑箱効果10: モンスター最大HP減少"""
        for m in self.state.monsters:
            m.max_hp = int(m.max_hp * 0.9)
        return "全モンスターの最大HPが少し減少した"

    def _green_box_effect_11(self):
        """緑箱効果11: 全箱HP1に"""
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                if land.condition <= LandCondition.PURPLE_BOX:
                    land.hp = 1
        return "全ての箱のHPの残りが1になった。"

    def _green_box_effect_12(self):
        """緑箱効果12: 赤箱消滅"""
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                if land.condition == LandCondition.RED_BOX:
                    land.condition = LandCondition.ROOM
        return "赤箱が消えた。"

    def _box_effect_purple(self, land: LandSquare):
        """紫箱の効果（コマンド使用回数増加）"""
        y = random.randint(1, 5 - self.state.difficulty) + 1

        ability_map = {
            (0, 2): (0, "全体攻撃"),
            (3, 6): (1, "HP全回復"),
            7: (2, "全体除去"),
            8: (3, "モンスター除去"),
            (9, 11): (4, "次の階へ行く効果"),
            (12, 14): (5, "復活の像"),
        }

        for key, (idx, name) in ability_map.items():
            if isinstance(key, tuple):
                if key[0] <= land.ability <= key[1]:
                    self.state.ability_hp[idx] += y
                    land.explanation = f"{name}の使用回数が{y}増えた。"
                    break
            else:
                if land.ability == key:
                    self.state.ability_hp[idx] += y
                    land.explanation = f"{name}の使用回数が{y}増えた。"
                    break

    # ========================================================================
    # ステータスチェック (Status Check)
    # ========================================================================

    def _status_check(self, level_up: bool):
        """ステータスチェックとレベルアップ処理"""
        p = self.state.player

        # プレイヤーのレベルアップ
        if p.exp >= p.level ** 3 and p.level < 999 and p.alive:
            dhp = random.randint(5, 7)
            p.level += 1
            p.max_hp += dhp
            p.hp += dhp
            p.atk = int(p.atk * 1.1) + 1
            p.defense += 1

        # ステータス上限チェック（プレイヤー）
        p.atk = min(p.atk, 10 ** 9 - 1)
        p.defense = max(min(p.defense, 10 ** 9 - 1), 1)
        p.exp = min(p.exp, 10 ** 9 - 1)
        p.max_hp = min(p.max_hp, 10 ** 7 - 1)
        p.hp = min(p.hp, p.max_hp)

        # ターン上限
        self.state.turn = min(self.state.turn, 10 ** 4 - 1)

        # ゲームオーバーチェック
        if p.hp <= 0 or self.state.turn <= 0:
            if self.state.ability_hp[5] == 0:
                p.hp = 0
                p.direction = 1
                p.alive = False
                self.state.game_mode = GameMode.GAME_OVER

        # モンスターのレベルアップ
        if level_up:
            for i in range(self.state.monster_number):
                m = self.state.monsters[i]
                if m.level < 999:
                    m.level += 1
                    m.max_hp = int(m.max_hp * 1.1) + 1
                    m.hp = m.max_hp
                    m.atk = int(m.atk * 1.2) + 1
                    m.defense += 1
                    m.exp += 10

                    if i == 0:
                        self.state.words[0] = f"ヘキサスライム0はLevel{m.level}になった。"

        # モンスターのステータス上限
        for m in self.state.monsters:
            m.atk = min(m.atk, 10 ** 9 - 1)
            m.defense = max(min(m.defense, 10 ** 9 - 1), 1)
            m.exp = min(m.exp, 10 ** 4 - 1)
            m.max_hp = min(m.max_hp, 10 ** 6 - 1)
            m.hp = min(m.hp, m.max_hp)

        # 地形のステータス上限
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                land.defense = max(min(land.defense, 10 ** 9 - 1), 1)
                land.max_hp = min(land.max_hp, 10 ** 6 - 1)
                land.hp = min(land.hp, land.max_hp)

    # ========================================================================
    # フロア設定 (Floor Set)
    # ========================================================================

    def _floor_set(self):
        """新しいフロアの設定"""
        self.state.floor += 1
        self.state.random_position.hp = GameState.LAND_NUMBER ** 2

        # 特定フロアでモンスター数増加
        if self.state.floor == 1:
            self._floor_set_initial()
        elif self.state.floor in [20, 40, 60, 80, 100]:
            self.state.monster_number += 10
        elif self.state.floor in [200, 300, 400, 500, 600, 700, 800]:
            self.state.monster_number += 100
        elif self.state.floor == 900:
            self.state.monster_number = 1000
        elif self.state.floor == 1000:
            self._floor_set_clear()
            return

        # マップ生成
        self._map_set()

        # プレイヤー配置
        self._place_player()

        # 階段配置
        self._place_stair()

        # 紫箱配置
        self._place_purple_box()

        # モンスター配置
        self._place_monsters()

        # 2階以降はモンスターレベルアップ
        if self.state.floor >= 2:
            self._status_check(True)

    def _floor_set_initial(self):
        """初期フロア設定"""
        self.state.words = [""] * 12

        p = self.state.player
        p.hp = 20
        p.max_hp = 20
        p.level = 1
        p.exp = 0
        p.height = GameState.SC
        p.width = GameState.SC
        p.left = 0
        p.top = 0
        p.atk = 10
        p.defense = 3
        p.direction = 1
        p.condition = 0

        self.state.turn = 200
        self.state.monster_number = 10

    def _floor_set_clear(self):
        """ゲームクリア処理"""
        p = self.state.player
        p.direction = 1
        p.alive = False
        self.state.game_mode = GameMode.GAME_CLEAR

    def _map_set(self):
        """マップ生成（簡易版）"""
        # 全て部屋で初期化
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                land.condition = LandCondition.ROOM
                land.alive = True
                land.hp = 10
                land.max_hp = 10
                land.defense = 1
                land.name = ""

        # 外周を壁に
        for i in range(GameState.LAND_NUMBER):
            self.state.land_squares[i][0].condition = LandCondition.WALL
            self.state.land_squares[i][GameState.LAND_NUMBER - 1].condition = LandCondition.WALL
            self.state.land_squares[0][i].condition = LandCondition.WALL
            self.state.land_squares[GameState.LAND_NUMBER - 1][i].condition = LandCondition.WALL

        # ランダムに壁と箱を配置
        for _ in range(100):
            x = random.randint(1, GameState.LAND_NUMBER - 2)
            y = random.randint(1, GameState.LAND_NUMBER - 2)
            rand = random.random()

            if rand < 0.1:
                self.state.land_squares[x][y].condition = LandCondition.WALL
            elif rand < 0.2:
                self.state.land_squares[x][y].condition = random.randint(0, 3)

    def _place_player(self):
        """プレイヤーを配置"""
        sc = GameState.SC
        ln = GameState.LAND_NUMBER

        while True:
            x = random.randint(1, ln - 2)
            y = random.randint(1, ln - 2)

            if self.state.land_squares[x][y].condition == LandCondition.ROOM:
                self.state.player.left = x * sc
                self.state.player.top = y * sc
                self.state.player.oleft = x * sc
                self.state.player.otop = y * sc
                self.state.player.alive = True
                self.state.player.condition = Ability.NO_ABILITY
                self.state.land_squares[x][y].alive = False
                break

    def _place_stair(self):
        """階段を配置"""
        sc = GameState.SC
        ln = GameState.LAND_NUMBER

        while True:
            x = random.randint(1, ln - 2)
            y = random.randint(1, ln - 2)

            if (self.state.land_squares[x][y].alive and
                self.state.land_squares[x][y].condition == LandCondition.ROOM):
                self.state.land_squares[x][y].condition = LandCondition.STAIR
                break

    def _place_purple_box(self):
        """紫箱を配置"""
        sc = GameState.SC
        ln = GameState.LAND_NUMBER
        attempts = 0

        while attempts < ln ** 2:
            x = random.randint(1, ln - 2)
            y = random.randint(1, ln - 2)

            if (self.state.land_squares[x][y].alive and
                self.state.land_squares[x][y].condition == LandCondition.ROOM):
                self.state.land_squares[x][y].condition = LandCondition.PURPLE_BOX
                break

            attempts += 1

    def _place_monsters(self):
        """モンスターを配置"""
        sc = GameState.SC
        ln = GameState.LAND_NUMBER

        for i in range(1000):
            m = self.state.monsters[i]

            if i >= self.state.monster_number:
                m.alive = False
                m.condition = 0
                continue

            # ランダムな位置を探す
            for _ in range(ln ** 2):
                x = random.randint(1, ln - 2)
                y = random.randint(1, ln - 2)

                if (self.state.land_squares[x][y].alive and
                    self.state.land_squares[x][y].condition == LandCondition.ROOM):
                    m.left = x * sc
                    m.top = y * sc
                    m.oleft = x * sc
                    m.otop = y * sc
                    m.alive = True
                    m.hp = m.max_hp
                    m.condition = 0
                    self.state.land_squares[x][y].condition = LandCondition.ENEMY
                    break
            else:
                m.alive = False
                m.condition = 0

    # ========================================================================
    # 特殊能力 (Abilities)
    # ========================================================================

    def ability_effect(self, key: str):
        """特殊能力の効果を適用"""
        if key == 'z':
            self._ability_all_attack()
        elif key == 'x':
            self._ability_hp_recovery()
        elif key == 'c':
            self._ability_all_removal()
        elif key == 'd':
            self._ability_monster_removal()
        elif key == 'enter':
            self._ability_next_floor()
        elif key == 'a':
            self._ability_green_box_trigger()

    def _ability_all_attack(self):
        """全体攻撃"""
        if self.state.ability_hp[0] > 0:
            self.state.ability_hp[0] -= 1

            p = self.state.player
            m0 = self.state.monsters[0]
            damage = int(p.atk * (random.random() * 0.2 + 0.9) / m0.defense)
            damage = min(damage, 10 ** 7 - 1)

            sc = GameState.SC
            for i in range(self.state.monster_number):
                m = self.state.monsters[i]
                if m.alive:
                    m.hp -= damage

                    if m.hp <= 0:
                        m.hp = 0
                        m.alive = False
                        x, y = m.left // sc, m.top // sc
                        self.state.land_squares[x][y].condition = LandCondition.ROOM
                        p.exp += m.exp

            self.state.words[2] = f"全モンスターは{damage}ダメージを受けた"
            p.direction = 11

    def _ability_hp_recovery(self):
        """HP全回復"""
        if self.state.ability_hp[1] > 0:
            self.state.ability_hp[1] -= 1
            self.state.player.hp = self.state.player.max_hp

    def _ability_all_removal(self):
        """全体除去"""
        if self.state.ability_hp[2] > 0:
            self.state.ability_hp[2] -= 1

            # 全箱・壁を除去
            for i in range(GameState.LAND_NUMBER):
                for j in range(GameState.LAND_NUMBER):
                    land = self.state.land_squares[i][j]
                    if land.condition <= LandCondition.GREEN_BOX or land.condition == LandCondition.WALL:
                        land.condition = LandCondition.ROOM

            # 全モンスター除去
            sc = GameState.SC
            for i in range(self.state.monster_number):
                m = self.state.monsters[i]
                if m.alive:
                    m.alive = False
                    m.hp = 0
                    m.condition = 0
                    x, y = m.left // sc, m.top // sc
                    self.state.land_squares[x][y].condition = LandCondition.ROOM

    def _ability_monster_removal(self):
        """モンスター除去"""
        if self.state.ability_hp[3] > 0 and self.state.player.alive:
            self.state.ability_hp[3] -= 1

            sc = GameState.SC
            for i in range(self.state.monster_number):
                m = self.state.monsters[i]
                if m.alive:
                    m.alive = False
                    m.hp = 0
                    m.condition = 0
                    x, y = m.left // sc, m.top // sc
                    self.state.land_squares[x][y].condition = random.randint(0, 3)

    def _ability_next_floor(self):
        """次のフロアへ"""
        if self.state.ability_hp[4] > 0:
            self.state.ability_hp[4] -= 1
            self._floor_set()

    def _ability_green_box_trigger(self):
        """緑箱の効果を全て発動"""
        for i in range(GameState.LAND_NUMBER):
            for j in range(GameState.LAND_NUMBER):
                land = self.state.land_squares[i][j]
                if land.condition == LandCondition.GREEN_BOX:
                    land.ability = random.randint(0, 14)
                    self.state.turn += 10
                    self._box_effect(i, j)

        self._floor_set()

    # ========================================================================
    # ゲーム制御 (Game Control)
    # ========================================================================

    def handle_key(self, key: str):
        """キー入力処理"""
        p = self.state.player

        if self.state.game_mode == GameMode.ENTRANCE:
            if key == 'enter':
                self.state.game_mode = GameMode.DUNGEON
                self.state.reset_dungeon()
                self.state.initialize_monsters()
                self._floor_set()

        elif self.state.game_mode == GameMode.DUNGEON:
            if key == 'up':
                p.direction *= 2
            elif key == 'down':
                p.direction *= 3
            elif key == 'right':
                p.direction *= 5
            elif key == 'left':
                p.direction *= 7
            elif key in ['z', 'x', 'c', 'd', 'enter', 'a']:
                self.ability_effect(key)

    def update(self):
        """ゲーム更新"""
        if self.state.game_mode == GameMode.DUNGEON:
            self.movement()
            self._status_check(False)


# ============================================================================
# メイン実行
# ============================================================================

if __name__ == "__main__":
    print("ダンジョンと不思議の箱 - ゲームエンジンモジュール")
    print("このモジュールは直接実行できません。")
    print("console_ui.py を使用してゲームをプレイしてください。")
