"""
Player Ability System
プレイヤーの特殊能力システム
"""
from models import GameState
from constants import TileType, LAND_NUMBER


class AbilitySystem:
    """プレイヤーのアビリティを管理するクラス"""

    def __init__(self, game_logic):
        self.game = game_logic

    def use_all_attack(self) -> str:
        """
        Z key: 全体攻撃
        全てのモンスターに一律ダメージを与える
        """
        if self.game.game_state.ability_hp[0] <= 0:
            return "使用回数が残っていない"

        self.game.game_state.ability_hp[0] -= 1

        damage, message = self.game.combat.all_monster_attack(
            self.game.player,
            self.game.monsters
        )

        # 死んだモンスターの位置をクリア
        for monster in self.game.monsters:
            if not monster.alive:
                self.game.map_gen.clear_monster(monster.left, monster.top)

        self.game.player.direction = 11  # スキップ用の特殊値

        return message

    def use_full_heal(self) -> str:
        """
        X key: HP全回復
        プレイヤーのHPを最大まで回復
        """
        if self.game.game_state.ability_hp[1] <= 0:
            return "使用回数が残っていない"

        self.game.game_state.ability_hp[1] -= 1
        self.game.player.hp = self.game.player.max_hp

        return "HPが全回復した"

    def use_clear_all(self) -> str:
        """
        C key: 全消去
        全ての箱、壁（階段以外）、モンスターを消去
        """
        if self.game.game_state.ability_hp[2] <= 0:
            return "使用回数が残っていない"

        self.game.game_state.ability_hp[2] -= 1

        # 箱と壁を消去
        for i in range(LAND_NUMBER):
            for j in range(LAND_NUMBER):
                square = self.game.map_gen.land_squares[i][j]
                if square.condition <= TileType.GREEN_BOX or square.condition == TileType.WALL:
                    square.condition = TileType.ROOM

        # 全モンスターを消去
        for monster in self.game.monsters:
            if monster.alive:
                monster.alive = False
                monster.hp = 0
                monster.condition = 0
                self.game.map_gen.clear_monster(monster.left, monster.top)

        return "全ての箱、壁、モンスターを消去した"

    def use_monster_removal(self) -> str:
        """
        D key: モンスター除去
        全てのモンスターを箱に変換
        """
        if self.game.game_state.ability_hp[3] <= 0:
            return "使用回数が残っていない"

        if not self.game.player.alive:
            return "プレイヤーが生きていない"

        self.game.game_state.ability_hp[3] -= 1

        import random

        # 全モンスターを箱に変換
        for monster in self.game.monsters:
            if monster.alive:
                monster.alive = False
                monster.hp = 0
                monster.condition = 0

                # モンスターがいた場所をランダムな箱に変換
                grid_x = monster.left // 10  # SC
                grid_y = monster.top // 10
                if 0 <= grid_x < LAND_NUMBER and 0 <= grid_y < LAND_NUMBER:
                    square = self.game.map_gen.land_squares[grid_x][grid_y]
                    square.condition = TileType(random.randint(0, 3))  # 青/赤/黄/緑箱

        return "全てのモンスターを箱に変換した"

    def use_next_floor(self) -> str:
        """
        Enter key: 次の階へ
        強制的に次のフロアへ進む
        """
        if self.game.game_state.ability_hp[4] <= 0:
            return "使用回数が残っていない"

        self.game.game_state.ability_hp[4] -= 1
        self.game.next_floor()

        return f"{self.game.game_state.floor}階に進んだ"

    def use_open_all_green_boxes(self) -> str:
        """
        A key: 緑箱を全て開ける
        フロア内の全ての緑箱の効果を発動してから次のフロアへ
        """
        count = 0

        for i in range(LAND_NUMBER):
            for j in range(LAND_NUMBER):
                square = self.game.map_gen.land_squares[i][j]

                if square.condition == TileType.GREEN_BOX:
                    import random
                    square.ability = random.randint(0, 14)

                    if self.game.player.condition != 3:  # TurnConst
                        self.game.game_state.turn += 10

                    self.game.box_effects.apply_box_effect(
                        self.game.player, square, self.game.game_state
                    )
                    count += 1

        if count > 0:
            self.game.next_floor()
            return f"{count}個の緑箱を開けて次のフロアへ進んだ"
        else:
            return "緑箱が見つからなかった"

    def use_resurrection(self) -> str:
        """
        復活の札（自動発動）
        ゲームオーバー時に自動で使用
        """
        if self.game.game_state.ability_hp[5] <= 0:
            return ""

        self.game.game_state.ability_hp[5] -= 1

        self.game.player.alive = True
        self.game.player.hp = self.game.player.max_hp

        if self.game.player.condition != 3:  # TurnConst
            self.game.game_state.turn += 100

        self.game.player.condition = 0  # Normal

        return "復活の札の効果で復活した。"
