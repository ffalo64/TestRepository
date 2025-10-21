"""
Map Generation System
マップ生成システム
"""
import random
from typing import List
from models import LandSquare, MonsterStatus
from constants import LAND_NUMBER, SC, TileType


class MapGenerator:
    """ダンジョンマップの生成を管理するクラス"""

    def __init__(self):
        self.land_squares: List[List[LandSquare]] = []
        self._initialize_map()

    def _initialize_map(self):
        """マップを初期化"""
        self.land_squares = [
            [LandSquare() for _ in range(LAND_NUMBER)]
            for _ in range(LAND_NUMBER)
        ]

        # 各マスの位置とサイズを設定
        for i in range(LAND_NUMBER):
            for j in range(LAND_NUMBER):
                square = self.land_squares[i][j]
                square.width = SC
                square.height = SC
                square.left = SC * i
                square.top = SC * j
                square.condition = TileType.ROOM
                square.alive = True

    def generate_dungeon(self, monster_max_hp: int, monster_def: int) -> int:
        """
        ダンジョンマップを生成
        Returns: 利用可能なマス数
        """
        # ランダムにマップパターンを生成
        # VB6版ではMap40.bmpから読み込んでいたが、Pythonでは乱数で生成
        available_spaces = 0

        for i in range(LAND_NUMBER):
            for j in range(LAND_NUMBER):
                square = self.land_squares[i][j]

                # 外周は必ず壁
                if i == 0 or j == 0 or i == LAND_NUMBER - 1 or j == LAND_NUMBER - 1:
                    square.condition = TileType.WALL
                    square.alive = True
                else:
                    # 内部はランダムに生成
                    rand = random.random()
                    if rand < 0.05:  # 5% 青箱
                        square.condition = TileType.BLUE_BOX
                        available_spaces += 1
                    elif rand < 0.10:  # 5% 赤箱
                        square.condition = TileType.RED_BOX
                        available_spaces += 1
                    elif rand < 0.15:  # 5% 黄箱
                        square.condition = TileType.YELLOW_BOX
                        available_spaces += 1
                    elif rand < 0.20:  # 5% 緑箱
                        square.condition = TileType.GREEN_BOX
                        available_spaces += 1
                    elif rand < 0.35:  # 15% 壁
                        square.condition = TileType.WALL
                    else:  # 65% 部屋
                        square.condition = TileType.ROOM
                        available_spaces += 1

                    square.alive = True

                # 箱と壁のHPと防御力を設定
                if square.condition <= TileType.PURPLE_BOX or square.condition == TileType.WALL:
                    square.max_hp = monster_max_hp
                    square.hp = square.max_hp
                    square.defense = monster_def

        return available_spaces

    def place_stairs(self) -> bool:
        """階段をランダムな位置に配置"""
        max_attempts = LAND_NUMBER ** 2
        for _ in range(max_attempts):
            x = random.randint(0, LAND_NUMBER - 1)
            y = random.randint(0, LAND_NUMBER - 1)

            square = self.land_squares[x][y]
            if square.alive and square.condition == TileType.ROOM:
                square.condition = TileType.STAIR
                return True

        return False

    def place_purple_box(self) -> bool:
        """紫箱をランダムな位置に配置"""
        max_attempts = LAND_NUMBER ** 2
        for _ in range(max_attempts):
            x = random.randint(0, LAND_NUMBER - 1)
            y = random.randint(0, LAND_NUMBER - 1)

            square = self.land_squares[x][y]
            if square.alive and square.condition == TileType.ROOM:
                square.condition = TileType.PURPLE_BOX
                return True

        return False

    def find_random_room(self) -> tuple[int, int] | None:
        """ランダムな部屋の位置を探す"""
        max_attempts = LAND_NUMBER ** 2
        for _ in range(max_attempts):
            x = random.randint(0, LAND_NUMBER - 1)
            y = random.randint(0, LAND_NUMBER - 1)

            square = self.land_squares[x][y]
            if square.alive and square.condition == TileType.ROOM:
                return (x, y)

        return None

    def place_monster(self, x: int, y: int):
        """指定位置にモンスターを配置"""
        square = self.land_squares[x][y]
        square.condition = TileType.ENEMY

    def clear_monster(self, x: int, y: int):
        """指定位置のモンスターを除去"""
        square = self.land_squares[x][y]
        if square.condition == TileType.ENEMY:
            square.condition = TileType.ROOM

    def get_square(self, x: int, y: int) -> LandSquare | None:
        """指定座標のマスを取得"""
        grid_x = x // SC
        grid_y = y // SC

        if 0 <= grid_x < LAND_NUMBER and 0 <= grid_y < LAND_NUMBER:
            return self.land_squares[grid_x][grid_y]
        return None

    def set_tile_name(self, x: int, y: int, name: str):
        """タイルの名前を設定"""
        square = self.get_square(x, y)
        if square:
            square.name = name

    def update_tile_names(self):
        """全タイルの名前を更新"""
        name_map = {
            TileType.BLUE_BOX: "青箱",
            TileType.RED_BOX: "赤箱",
            TileType.YELLOW_BOX: "黄箱",
            TileType.GREEN_BOX: "緑箱",
            TileType.PURPLE_BOX: "紫箱",
            TileType.WALL: "壁",
        }

        for i in range(LAND_NUMBER):
            for j in range(LAND_NUMBER):
                square = self.land_squares[i][j]
                square.name = name_map.get(square.condition, "")

    def clear_walls_except_border(self):
        """外周以外の壁を全て消去"""
        for i in range(1, LAND_NUMBER - 1):
            for j in range(1, LAND_NUMBER - 1):
                square = self.land_squares[i][j]
                if square.condition == TileType.WALL:
                    square.condition = TileType.ROOM

    def convert_all_boxes(self, new_type: TileType):
        """全ての箱を指定タイプに変換"""
        for i in range(LAND_NUMBER):
            for j in range(LAND_NUMBER):
                square = self.land_squares[i][j]
                if square.condition <= TileType.GREEN_BOX:
                    square.condition = new_type

    def destroy_all_boxes(self):
        """全ての箱を消去"""
        for i in range(LAND_NUMBER):
            for j in range(LAND_NUMBER):
                square = self.land_squares[i][j]
                if square.condition <= TileType.GREEN_BOX:
                    square.condition = TileType.ROOM

    def destroy_red_boxes(self):
        """赤箱を全て消去"""
        for i in range(LAND_NUMBER):
            for j in range(LAND_NUMBER):
                square = self.land_squares[i][j]
                if square.condition == TileType.RED_BOX:
                    square.condition = TileType.ROOM
