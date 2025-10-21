"""
Combat System
戦闘システム
"""
import random
from models import PlayerStatus, MonsterStatus, LandSquare
from constants import MAX_DAMAGE, LAND_NUMBER, SC, TileType, PlayerCondition


class CombatSystem:
    """戦闘とダメージ計算を管理するクラス"""

    @staticmethod
    def calculate_damage(attacker_atk: int, defender_def: int) -> int:
        """
        ダメージを計算
        Formula: ATK * (0.9~1.1 random) / DEF
        """
        if defender_def == 0:
            defender_def = 1

        # ランダム係数: 0.9~1.1
        random_factor = random.uniform(0.9, 1.1)
        damage = int(attacker_atk * random_factor / defender_def) + 1

        # ダメージ上限
        if damage > MAX_DAMAGE:
            damage = MAX_DAMAGE

        return damage

    def player_attack_monster(
        self,
        player: PlayerStatus,
        monster: MonsterStatus,
        map_gen
    ) -> tuple[int, str]:
        """
        プレイヤーがモンスターを攻撃
        Returns: (damage, message)
        """
        if not monster.alive:
            return (0, "")

        damage = self.calculate_damage(player.atk, monster.defense)
        monster.hp -= damage

        message = f"{monster.name}に{damage}ダメージを与えた"

        if monster.hp <= 0:
            monster.hp = 0
            monster.alive = False

            # モンスターがいた場所を部屋に戻す
            grid_x = monster.left // SC
            grid_y = monster.top // SC
            map_gen.clear_monster(monster.left, monster.top)

            # 経験値を獲得
            player.exp += monster.exp

        return (damage, message)

    def player_attack_tile(
        self,
        player: PlayerStatus,
        tile: LandSquare,
        map_gen
    ) -> tuple[int, str]:
        """
        プレイヤーが箱や壁を攻撃
        Returns: (damage, message)
        """
        damage = self.calculate_damage(player.atk, tile.defense)
        tile.hp -= damage

        message = f"{tile.name}に{damage}ダメージを与えた"

        if tile.hp <= 0:
            tile.hp = 0
            # 箱を破壊した時の処理は box_effects で行う
            tile.ability = random.randint(0, 14)  # ランダム効果

        return (damage, message)

    def monster_attack_player(
        self,
        monster: MonsterStatus,
        player: PlayerStatus
    ) -> int:
        """
        モンスターがプレイヤーを攻撃
        Returns: damage
        """
        damage = self.calculate_damage(monster.atk, player.defense)
        player.hp -= damage
        return damage

    def monster_attack_tile(
        self,
        monster: MonsterStatus,
        tile: LandSquare
    ) -> int:
        """
        モンスターが箱や壁を攻撃
        Returns: damage
        """
        if tile.condition == TileType.STAIR:
            return 0  # 階段は攻撃できない

        damage = self.calculate_damage(monster.atk, tile.defense)
        tile.hp -= damage

        if tile.hp <= 0:
            tile.hp = 0
            tile.condition = TileType.ROOM

        return damage

    def all_monster_attack(
        self,
        player: PlayerStatus,
        monsters: list[MonsterStatus]
    ) -> tuple[int, str]:
        """
        全モンスターに一律ダメージを与える
        Returns: (damage, message)
        """
        # 最初のモンスターの防御力を基準にダメージ計算
        if not monsters or not monsters[0]:
            return (0, "")

        damage = self.calculate_damage(player.atk, monsters[0].defense)
        killed_count = 0

        for monster in monsters:
            if monster.alive:
                monster.hp -= damage

                if monster.hp <= 0:
                    monster.hp = 0
                    monster.alive = False
                    killed_count += 1
                    player.exp += monster.exp

        message = f"全てのモンスターに{damage}ダメージを与えた"
        return (damage, message)
