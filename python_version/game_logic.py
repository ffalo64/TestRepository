"""
Core Game Logic
ゲームロジック全体の統合
"""
import random
from models import PlayerStatus, MonsterStatus, GameState
from constants import (
    LAND_NUMBER, SC, PLAYER_CONST, MAX_MONSTERS, INITIAL_TURN, GOAL_FLOOR,
    TileType, MonsterAbility, PlayerCondition, GameMode,
    MAX_ATK, MAX_DEF, MAX_EXP, MAX_PLAYER_HP, MAX_MONSTER_HP, MAX_TURN
)
from map_generator import MapGenerator
from combat import CombatSystem
from box_effects import BoxEffects


class GameLogic:
    """ゲームロジックを管理するクラス"""

    def __init__(self):
        self.player = PlayerStatus()
        self.monsters: list[MonsterStatus] = []
        self.game_state = GameState()
        self.map_gen = MapGenerator()
        self.combat = CombatSystem()
        self.box_effects: BoxEffects | None = None

        self._initialize_monsters()
        self.box_effects = BoxEffects(self.map_gen, self.monsters)

    def _initialize_monsters(self):
        """モンスターを初期化"""
        self.monsters = []
        for i in range(MAX_MONSTERS):
            monster = MonsterStatus()
            monster.name = f"ヘキサスライム{i}"
            monster.hp = 10
            monster.max_hp = 10
            monster.level = 1
            monster.exp = 10
            monster.ability = MonsterAbility.NO_ABILITY
            monster.atk = 10
            monster.defense = 1
            monster.width = SC
            monster.height = SC
            monster.direction = 1
            monster.alive = False
            self.monsters.append(monster)

    def start_new_game(self):
        """新しいゲームを開始"""
        self.player.reset()
        self.game_state.reset()
        self.game_state.game_mode = GameMode.DUNGEON
        self._setup_floor(is_first=True)

    def _setup_floor(self, is_first: bool = False):
        """フロアをセットアップ"""
        self.game_state.floor += 1

        if is_first:
            # 初回設定
            self.player.hp = 20
            self.player.max_hp = 20
            self.player.level = 1
            self.player.exp = 0
            self.player.atk = 10
            self.player.defense = 3
            self.player.condition = PlayerCondition.NORMAL
            self.game_state.turn = INITIAL_TURN
            self.game_state.monster_number = 10

        # マップ生成
        available_spaces = self.map_gen.generate_dungeon(
            self.monsters[0].max_hp,
            self.monsters[0].defense
        )

        # プレイヤーを配置
        self._place_player()

        # 階段を配置
        self.map_gen.place_stairs()

        # 紫箱を配置
        self.map_gen.place_purple_box()

        # モンスターを配置
        self._place_monsters()

        # 特定階層でモンスター数を増やす
        if self.game_state.floor in [20, 40, 60, 80, 100]:
            self.game_state.monster_number += 10
        elif self.game_state.floor in [200, 300, 400, 500, 600, 700, 800]:
            self.game_state.monster_number += 100
        elif self.game_state.floor == 900:
            self.game_state.monster_number = 1000

        # 2階以降はモンスターレベルアップ
        if self.game_state.floor >= 2:
            self._level_up_monsters()

        # タイル名を更新
        self.map_gen.update_tile_names()

        # プレイヤーの状態異常をリセット
        if self.player.condition != PlayerCondition.NORMAL:
            self.player.condition = PlayerCondition.NORMAL

    def _place_player(self):
        """プレイヤーをランダムな部屋に配置"""
        for _ in range(LAND_NUMBER ** 2):
            x = random.randint(0, LAND_NUMBER - 1)
            y = random.randint(0, LAND_NUMBER - 1)

            square = self.map_gen.land_squares[x][y]
            if square.alive and square.condition == TileType.ROOM:
                self.player.left = x * SC
                self.player.top = y * SC
                self.player.oleft = self.player.left
                self.player.otop = self.player.top
                self.player.alive = True
                square.alive = False
                return

    def _place_monsters(self):
        """モンスターをランダムに配置"""
        placed = 0
        for monster in self.monsters:
            if placed >= self.game_state.monster_number:
                monster.alive = False
                break

            for _ in range(LAND_NUMBER ** 2):
                x = random.randint(0, LAND_NUMBER - 1)
                y = random.randint(0, LAND_NUMBER - 1)

                square = self.map_gen.land_squares[x][y]
                if square.alive and square.condition == TileType.ROOM:
                    monster.left = x * SC
                    monster.top = y * SC
                    monster.oleft = monster.left
                    monster.otop = monster.top
                    monster.alive = True
                    monster.hp = monster.max_hp
                    square.condition = TileType.ENEMY
                    placed += 1
                    break
            else:
                monster.alive = False

    def _level_up_monsters(self):
        """全モンスターのレベルアップ"""
        for monster in self.monsters:
            if monster.level < 999:
                monster.level += 1
                monster.max_hp = int(monster.max_hp * 1.1) + 1
                monster.hp = monster.max_hp
                monster.atk = int(monster.atk * 1.2) + 1
                monster.defense += 1
                monster.exp += 10

    def move_player(self, direction: int):
        """
        プレイヤーを移動
        direction: 2=上, 3=下, 5=右, 7=左
        """
        if not self.player.alive or self.player.direction == 1:
            return

        # 移動先を計算
        new_left = self.player.left
        new_top = self.player.top

        if direction % 2 == 0 and self.player.top > 0:  # 上
            new_top -= SC
        if direction % 3 == 0 and self.player.top < (LAND_NUMBER - 1) * SC:  # 下
            new_top += SC
        if direction % 5 == 0 and self.player.left < (LAND_NUMBER - 1) * SC:  # 右
            new_left += SC
        if direction % 7 == 0 and self.player.left > 0:  # 左
            new_left -= SC

        # 移動先のチェック
        self.player.left = new_left
        self.player.top = new_top
        self._check_player_position()

        # ターン減少
        if self.player.condition != PlayerCondition.SLOW:
            self.game_state.turn -= 1
        elif self.player.condition == PlayerCondition.SLOW and self.game_state.turn % 2 == 0:
            self.game_state.turn -= 1

        # モンスター移動
        self._move_monsters()

        # ステータスチェック
        self.check_status()

    def _check_player_position(self):
        """プレイヤーの位置をチェックして衝突判定"""
        square = self.map_gen.get_square(self.player.left, self.player.top)
        if not square:
            return

        # 階段
        if square.condition == TileType.STAIR:
            self._setup_floor()
            return

        # 箱
        if square.condition <= TileType.PURPLE_BOX:
            # 箱を攻撃
            damage, message = self.combat.player_attack_tile(
                self.player, square, self.map_gen
            )

            if square.hp <= 0:
                # 箱が壊れた - 効果を適用
                if self.player.condition != PlayerCondition.TURN_CONST:
                    self.game_state.turn += 10

                explanation = self.box_effects.apply_box_effect(
                    self.player, square, self.game_state
                )
                square.explanation = explanation
                square.ability = 0
                square.condition = TileType.ROOM

            # 移動キャンセル
            self.player.left = self.player.oleft
            self.player.top = self.player.otop
            return

        # 壁
        if square.condition == TileType.WALL:
            can_break = (
                self.player.condition == PlayerCondition.WALL_BREAK and
                self.player.left != 0 and self.player.top != 0 and
                self.player.left != (LAND_NUMBER - 1) * SC and
                self.player.top != (LAND_NUMBER - 1) * SC
            )

            if can_break:
                self.combat.player_attack_tile(self.player, square, self.map_gen)

            self.player.left = self.player.oleft
            self.player.top = self.player.otop
            return

        # モンスターとの衝突チェック
        for monster in self.monsters:
            if (monster.alive and
                monster.left == self.player.left and
                monster.top == self.player.top):

                damage, message = self.combat.player_attack_monster(
                    self.player, monster, self.map_gen
                )

                self.player.left = self.player.oleft
                self.player.top = self.player.otop
                return

        # 位置を確定
        self.player.oleft = self.player.left
        self.player.otop = self.player.top

    def _move_monsters(self):
        """全モンスターを移動"""
        for monster in self.monsters:
            if not monster.alive:
                continue

            # スロー判定
            if monster.ability == MonsterAbility.SLOW and self.game_state.turn % 2 != 0:
                continue

            # ステルス状態のプレイヤーは追跡しない
            if self.player.condition == PlayerCondition.STEALTH:
                continue

            # プレイヤーに向かって移動
            monster.direction = 1

            if monster.top > self.player.top:
                monster.top -= SC
                monster.direction *= 2
            if monster.top < self.player.top:
                monster.top += SC
                monster.direction *= 3
            if monster.left < self.player.left:
                monster.left += SC
                monster.direction *= 5
            if monster.left > self.player.left:
                monster.left -= SC
                monster.direction *= 7

            self._check_monster_position(monster)
            monster.direction = 1

    def _check_monster_position(self, monster: MonsterStatus):
        """モンスターの位置をチェック"""
        # プレイヤーとの衝突
        if (monster.left == self.player.left and
            monster.top == self.player.top):

            monster.left = monster.oleft
            monster.top = monster.otop

            damage = self.combat.monster_attack_player(monster, self.player)
            self.game_state.sum_damage += damage
            return

        square = self.map_gen.get_square(monster.left, monster.top)
        if not square:
            return

        # 部屋以外の場所
        if square.condition != TileType.ROOM:
            # 壁破壊
            can_break_wall = (
                square.condition == TileType.WALL and
                monster.ability == MonsterAbility.WALL_BREAK and
                monster.left != 0 and monster.top != 0 and
                monster.left != (LAND_NUMBER - 1) * SC and
                monster.top != (LAND_NUMBER - 1) * SC
            )

            # 箱攻撃
            can_attack_box = (
                square.condition <= TileType.PURPLE_BOX and
                monster.ability == MonsterAbility.BOX_ATTACK
            )

            if can_break_wall or can_attack_box:
                self.combat.monster_attack_tile(monster, square)

            # 迂回路を探す
            self._find_alternative_path(monster)
            return

        # 以前の位置を部屋に戻す
        old_square = self.map_gen.get_square(monster.oleft, monster.otop)
        if old_square:
            old_square.condition = TileType.ROOM

        # 新しい位置を敵に設定
        square.condition = TileType.ENEMY
        monster.oleft = monster.left
        monster.otop = monster.top

    def _find_alternative_path(self, monster: MonsterStatus):
        """モンスターの迂回路を探す"""
        # VB6の複雑なロジックを簡略化
        # 元の位置に戻す
        monster.left = monster.oleft
        monster.top = monster.otop

    def check_status(self):
        """ステータスをチェック"""
        # プレイヤーのレベルアップ
        if self.player.exp >= self.player.level ** 3 and self.player.level < 999:
            dhp = random.randint(5, 7)
            self.player.level += 1
            self.player.max_hp += dhp
            self.player.hp += dhp
            self.player.atk = int(self.player.atk * 1.1) + 1
            self.player.defense += 1

        # ステータス上限チェック
        if self.player.atk > MAX_ATK:
            self.player.atk = MAX_ATK
        if self.player.defense > MAX_DEF:
            self.player.defense = MAX_DEF
        if self.player.defense == 0:
            self.player.defense = 1
        if self.game_state.turn > MAX_TURN:
            self.game_state.turn = MAX_TURN
        if self.player.exp > MAX_EXP:
            self.player.exp = MAX_EXP
        if self.player.max_hp > MAX_PLAYER_HP:
            self.player.max_hp = MAX_PLAYER_HP

        # HP上限チェック
        if self.player.hp > self.player.max_hp:
            self.player.hp = self.player.max_hp

        # ゲームオーバーチェック
        if self.player.hp <= 0 or self.game_state.turn <= 0:
            if self.game_state.ability_hp[5] == 0:
                self.player.hp = 0
                self.player.alive = False
                self.game_state.game_mode = GameMode.GAME_OVER
            else:
                # 復活の札を使用
                self.game_state.ability_hp[5] -= 1
                self.player.alive = True
                self.player.hp = self.player.max_hp
                if self.player.condition != PlayerCondition.TURN_CONST:
                    self.game_state.turn += 100
                self.player.condition = PlayerCondition.NORMAL

        # ゴールチェック
        if self.game_state.floor >= GOAL_FLOOR:
            self.player.alive = False
            self.game_state.game_mode = GameMode.GAME_CLEAR

        # モンスターのステータスチェック
        for monster in self.monsters:
            if monster.atk > MAX_ATK:
                monster.atk = MAX_ATK
            if monster.defense > MAX_DEF:
                monster.defense = MAX_DEF
            if monster.defense == 0:
                monster.defense = 1
            if monster.hp > monster.max_hp:
                monster.hp = monster.max_hp
            if monster.max_hp > MAX_MONSTER_HP:
                monster.max_hp = MAX_MONSTER_HP

    def next_floor(self):
        """次のフロアへ進む"""
        self._setup_floor()
