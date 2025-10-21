"""
Box Effects System
箱の効果システム
"""
import random
from models import PlayerStatus, MonsterStatus, LandSquare, GameState
from constants import TileType, LAND_NUMBER, SC, MonsterAbility, PlayerCondition


class BoxEffects:
    """箱を開けた時の効果を管理するクラス"""

    def __init__(self, map_gen, monsters: list[MonsterStatus]):
        self.map_gen = map_gen
        self.monsters = monsters

    def apply_box_effect(
        self,
        player: PlayerStatus,
        tile: LandSquare,
        game_state: GameState
    ) -> str:
        """
        箱の効果を適用
        Returns: 効果の説明文
        """
        if tile.condition == TileType.BLUE_BOX:
            return self._blue_box_effect(player, tile)
        elif tile.condition == TileType.RED_BOX:
            return self._red_box_effect(player, tile, game_state)
        elif tile.condition == TileType.YELLOW_BOX:
            return self._yellow_box_effect(player, tile, game_state)
        elif tile.condition == TileType.GREEN_BOX:
            return self._green_box_effect(player, tile)
        elif tile.condition == TileType.PURPLE_BOX:
            return self._purple_box_effect(player, tile, game_state)
        elif tile.condition == TileType.WALL:
            return ""

        return ""

    def _blue_box_effect(self, player: PlayerStatus, tile: LandSquare) -> str:
        """青箱の効果: HP小回復、モンスターの状態異常解除"""
        player.hp += player.max_hp // 10

        # 全モンスターの能力をリセット
        for monster in self.monsters:
            monster.ability = MonsterAbility.NO_ABILITY

        return "HPが少し回復し、全モンスターの状態異常が解除された。"

    def _red_box_effect(
        self,
        player: PlayerStatus,
        tile: LandSquare,
        game_state: GameState
    ) -> str:
        """赤箱の効果: 15種のマイナス効果（ランダム）"""
        effect = tile.ability

        if effect == 0:
            player.hp = 1
            return "HPが1になってしまった"

        elif effect == 1:
            player.atk = int(player.atk * 0.9)
            return "攻撃力が下がった"

        elif effect == 2:
            player.condition = PlayerCondition.SLOW
            return "プレイヤーはこのフロアにおいて、動きが鈍くなった。"

        elif effect == 3:
            player.condition = PlayerCondition.TURN_CONST
            return "このフロアにおいて、ターンが減らなくなった。"

        elif effect == 4:
            player.max_hp = int(player.max_hp * 0.9)
            return "最大HPが下がった"

        elif effect == 5:
            player.hp = int(player.hp / 2) + 1
            return "HPが半分になった"

        elif effect == 6:
            for monster in self.monsters:
                if monster.alive:
                    monster.hp = monster.max_hp
                    monster.condition = 0
            return "全モンスターが全回復した。"

        elif effect == 7:
            for monster in self.monsters:
                monster.atk = int(monster.atk * 1.1)
            return "全てのモンスターの攻撃力が上がった"

        elif effect == 8:
            # レベルアップ処理は game_logic で実行
            return "全てのモンスターのレベルが上がった"

        elif effect == 9:
            for i in range(LAND_NUMBER):
                for j in range(LAND_NUMBER):
                    square = self.map_gen.land_squares[i][j]
                    if square.condition <= TileType.PURPLE_BOX:
                        square.defense *= 2
            return "このフロアの全ての箱の守備力が2倍になった。"

        elif effect == 10:
            for monster in self.monsters:
                monster.defense = int(monster.defense * 1.1)
            return "全てのモンスターの守備力が上がった"

        elif effect == 11:
            self.map_gen.convert_all_boxes(TileType.RED_BOX)
            return "全ての箱が赤色になった"

        elif effect == 12:
            for monster in self.monsters:
                monster.max_hp = int(monster.max_hp * 1.1)
            return "全てのモンスターの最大HPが上がった"

        elif effect == 13:
            for monster in self.monsters:
                monster.exp = int(monster.exp * 0.9)
            return "全てのモンスターの経験値が少し減少した"

        elif effect == 14:
            for i in range(LAND_NUMBER):
                for j in range(LAND_NUMBER):
                    square = self.map_gen.land_squares[i][j]
                    if square.condition <= TileType.PURPLE_BOX:
                        square.hp *= 2
                        square.max_hp *= 2
            return "このフロアの全ての箱のHPが2倍になった。"

        return ""

    def _yellow_box_effect(
        self,
        player: PlayerStatus,
        tile: LandSquare,
        game_state: GameState
    ) -> str:
        """黄箱の効果: 15種の混合効果（ランダム）"""
        effect = tile.ability

        if effect == 0:
            for monster in self.monsters:
                monster.ability = MonsterAbility.STEALTH
            return "全モンスターが透明になった。"

        elif effect == 1:
            for i in range(len(game_state.ability_hp) - 1):
                game_state.ability_hp[i] += 1
            return "全てのコマンドの使用回数が1増えた。"

        elif effect == 2:
            player.condition = PlayerCondition.TURN_CONST
            return "このフロアにおいて、ターンが減らなくなった。"

        elif effect == 3:
            for monster in self.monsters:
                if monster.alive:
                    monster.alive = False
                    monster.hp = 0
                    monster.condition = 0
                    grid_x = monster.left // SC
                    grid_y = monster.top // SC
                    self.map_gen.clear_monster(monster.left, monster.top)
            return "全モンスターが全滅した。"

        elif effect == 4:
            player.defense *= 2
            return "守備力が2倍になった"

        elif effect == 5:
            player.atk *= 2
            return "攻撃力が2倍になった"

        elif effect == 6:
            game_state.turn = 1000
            return "ターンの残りが1000になった。"

        elif effect == 7:
            player.atk //= 2
            return "攻撃力が半分になった"

        elif effect == 8:
            for monster in self.monsters:
                monster.ability = MonsterAbility.WALL_BREAK
            return "全モンスターが壁を破壊するようになった。(画面外を除く)"

        elif effect == 9:
            for i in range(len(game_state.ability_hp) - 1):
                game_state.ability_hp[i] = 5
            return "全てのコマンドの使用回数が5になった。"

        elif effect == 10:
            self.map_gen.destroy_all_boxes()
            return "全ての箱が消えた。"

        elif effect == 11:
            self.map_gen.convert_all_boxes(TileType.BLUE_BOX)
            return "全ての箱が青色になった"

        elif effect == 12:
            for i in range(LAND_NUMBER):
                for j in range(LAND_NUMBER):
                    square = self.map_gen.land_squares[i][j]
                    if square.condition <= TileType.GREEN_BOX:
                        square.condition = TileType(random.randint(0, 3))
            return "全ての箱がランダムに変化した。"

        elif effect == 13:
            self.map_gen.convert_all_boxes(TileType.YELLOW_BOX)
            return "全ての箱が黄色になった"

        elif effect == 14:
            for monster in self.monsters:
                monster.ability = MonsterAbility.BOX_ATTACK
            return "全モンスターが箱を壊せるようになった。"

        return ""

    def _green_box_effect(self, player: PlayerStatus, tile: LandSquare) -> str:
        """緑箱の効果: 15種のプラス効果（ランダム）"""
        effect = tile.ability

        if effect == 0:
            player.hp = player.max_hp
            return "HPが全回復"

        elif effect == 1:
            player.atk = int(player.atk * 1.1) + 1
            return "攻撃力が上がった"

        elif effect == 2:
            for monster in self.monsters:
                if monster.alive:
                    monster.alive = False
                    monster.hp = 0
                    monster.condition = 0
                    grid_x = monster.left // SC
                    grid_y = monster.top // SC
                    self.map_gen.clear_monster(monster.left, monster.top)
            return "全モンスターが全滅した。"

        elif effect == 3:
            for monster in self.monsters:
                monster.ability = MonsterAbility.SLOW
            return "全モンスターの動きが遅くなった。"

        elif effect == 4:
            player.max_hp = int(player.max_hp * 1.1) + 1
            return "最大HPが上がった"

        elif effect == 5:
            self.map_gen.clear_walls_except_border()
            return "画面外以外の壁が全て壊れた。"

        elif effect == 6:
            player.condition = PlayerCondition.STEALTH
            return "プレイヤーは透明になった。(このフロアのみ)"

        elif effect == 7:
            for monster in self.monsters:
                monster.hp = 1
            return "全てのモンスターのHPが残り1になった"

        elif effect == 8:
            for monster in self.monsters:
                monster.atk = int(monster.atk * 0.9)
            return "全てのモンスターの攻撃力が少し減少した"

        elif effect == 9:
            for monster in self.monsters:
                monster.exp += 5
            return "全てのモンスターの経験値が少し上がった"

        elif effect == 10:
            for monster in self.monsters:
                monster.max_hp = int(monster.max_hp * 0.9)
            return "全てのモンスターの最大HPが少し減少した"

        elif effect == 11:
            for i in range(LAND_NUMBER):
                for j in range(LAND_NUMBER):
                    square = self.map_gen.land_squares[i][j]
                    if square.condition <= TileType.PURPLE_BOX:
                        square.hp = 1
            return "全ての箱のHPが残り1になった。"

        elif effect == 12:
            self.map_gen.destroy_red_boxes()
            return "赤箱が消えた。"

        elif effect == 13:
            player.defense += 1
            return "守備力が少し上がった"

        elif effect == 14:
            player.condition = PlayerCondition.WALL_BREAK
            return "このフロアにおいて、壁を破壊できるようになった(画面外を除く)"

        return ""

    def _purple_box_effect(
        self,
        player: PlayerStatus,
        tile: LandSquare,
        game_state: GameState
    ) -> str:
        """紫箱の効果: アビリティチャージ増加"""
        # 難易度によってチャージ量が変わる
        charge = random.randint(1, 5 - game_state.difficulty) + 1

        effect = tile.ability

        if effect <= 2:  # 0-2
            game_state.ability_hp[0] += charge
            return f"全の攻撃の使用回数が{charge}増えた。"

        elif effect <= 6:  # 3-6
            game_state.ability_hp[1] += charge
            return f"HP全快の使用回数が{charge}増えた。"

        elif effect == 7:
            game_state.ability_hp[2] += charge
            return f"全消去の使用回数が{charge}増えた。"

        elif effect == 8:
            game_state.ability_hp[3] += charge
            return f"モンスター除去の使用回数が{charge}増えた。"

        elif effect <= 11:  # 9-11
            game_state.ability_hp[4] += charge
            return f"次の階へ行く命令の使用回数が{charge}増えた。"

        else:  # 12-14
            game_state.ability_hp[5] += charge
            return f"復活の札が{charge}枚に増加した。"
