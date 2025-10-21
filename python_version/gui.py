"""
Tkinter-based GUI
tkinterベースのグラフィカルユーザーインターフェース
"""
import tkinter as tk
from tkinter import font as tkfont
from PIL import Image, ImageTk
import os
from game_logic import GameLogic
from ability_system import AbilitySystem
from constants import (
    LAND_NUMBER, SC, GameMode, TileType, MonsterAbility,
    PlayerCondition, Color
)


class DungeonGameGUI:
    """ダンジョンゲームのGUIクラス"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("ダンジョンと不思議の箱 - Python Edition")

        # ウィンドウサイズ: 600x600 ピクセル
        self.window_width = 600
        self.window_height = 600
        self.root.geometry(f"{self.window_width}x{self.window_height}")
        self.root.resizable(False, False)

        # ゲームロジック
        self.game = GameLogic()
        self.ability_system = AbilitySystem(self.game)

        # キャンバス
        self.canvas = tk.Canvas(
            self.root,
            width=self.window_width,
            height=self.window_height,
            bg='black',
            highlightthickness=0
        )
        self.canvas.pack()

        # フォント
        self.font = tkfont.Font(family="Meiryo", size=12, weight="bold")
        self.small_font = tkfont.Font(family="Meiryo", size=10)

        # 画像読み込み
        self.images = {}
        self._load_images()

        # メッセージバッファ
        self.messages: list[tuple[str, int]] = []  # (message, ttl)

        # キーバインディング
        self.root.bind("<KeyPress>", self.on_key_press)

        # ゲーム開始
        self.game.game_state.game_mode = GameMode.ENTRANCE

        # メインループ
        self.update_game()

    def _load_images(self):
        """画像ファイルを読み込む"""
        # プロジェクトルートからの相対パス
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        image_dir = os.path.join(base_path, "Images")

        try:
            # プレイヤー画像 (20x50 = 5状態 x 10px)
            player_img = Image.open(os.path.join(image_dir, "Player.bmp"))
            self.images['player'] = ImageTk.PhotoImage(player_img)
            self.images['player_pil'] = player_img

            # モンスター画像 (20x50 = 5種類 x 10px)
            monster_img = Image.open(os.path.join(image_dir, "Monster.bmp"))
            self.images['monster'] = ImageTk.PhotoImage(monster_img)
            self.images['monster_pil'] = monster_img

            # 箱画像 (20x60 = 6種類 x 10px: 青/赤/黄/緑/紫/階段)
            box_img = Image.open(os.path.join(image_dir, "Box.bmp"))
            self.images['box'] = ImageTk.PhotoImage(box_img)
            self.images['box_pil'] = box_img

            # 部屋画像 (20x100 = 10種類 x 10px)
            room_img = Image.open(os.path.join(image_dir, "Room.bmp"))
            self.images['room'] = ImageTk.PhotoImage(room_img)
            self.images['room_pil'] = room_img

            # 壁画像 (20x100 = 10種類 x 10px)
            wall_img = Image.open(os.path.join(image_dir, "Wall.bmp"))
            self.images['wall'] = ImageTk.PhotoImage(wall_img)
            self.images['wall_pil'] = wall_img

            # HPバー画像
            hpbar_img = Image.open(os.path.join(image_dir, "HpBar.bmp"))
            self.images['hpbar'] = ImageTk.PhotoImage(hpbar_img)
            self.images['hpbar_pil'] = hpbar_img

            print("画像の読み込みに成功しました")

        except Exception as e:
            print(f"警告: 画像の読み込みに失敗しました: {e}")
            print("代わりに色付き図形を使用します")
            self.images = {}  # 画像なしモード

    def _get_sprite(self, image_name: str, index: int) -> ImageTk.PhotoImage | None:
        """
        スプライトシートから特定のスプライトを切り出す（透過処理付き）
        VB6のBitBlt（vbSrcAnd + vbSrcPaint）を再現
        Args:
            image_name: 画像名 ('player', 'monster', 'box', 'room', 'wall')
            index: スプライトのインデックス（縦方向の位置）
        """
        if f'{image_name}_pil' not in self.images:
            return None

        pil_img = self.images[f'{image_name}_pil']

        # VB6では横20ピクセル（左10px=マスク、右10px=実画像）
        # BitBltのAND/OR演算を再現して透過処理を行う
        try:
            # 左10px: マスク（黒=透明、白=不透明）
            mask_img = pil_img.crop((0, index * SC, SC, (index + 1) * SC))

            # 右10px: 実画像
            sprite_img = pil_img.crop((10, index * SC, 20, (index + 1) * SC))

            # RGBAモードに変換
            if sprite_img.mode != 'RGBA':
                sprite_img = sprite_img.convert('RGBA')

            # マスクをグレースケールに変換
            if mask_img.mode != 'L':
                mask_img = mask_img.convert('L')

            # VB6のBitBltではマスクの黒い部分(0)が透明
            # Pillowのアルファチャンネルでは255=不透明、0=透明
            # そのため、マスクを反転する必要がある
            from PIL import ImageOps
            mask_inverted = ImageOps.invert(mask_img)

            # スプライト画像のピクセルデータを取得してアルファチャンネルを作成
            sprite_data = sprite_img.getdata()
            mask_data = list(mask_inverted.getdata())

            # 新しいピクセルデータ（RGBA）
            new_data = []
            for i, pixel in enumerate(sprite_data):
                r, g, b = pixel[:3] if len(pixel) >= 3 else (pixel[0], pixel[0], pixel[0])

                # マスクのアルファ値を取得
                alpha = mask_data[i]

                # 実画像が黒(0,0,0)の場合も透明にする（背景色）
                if r == 0 and g == 0 and b == 0:
                    alpha = 0

                new_data.append((r, g, b, alpha))

            # 新しい画像を作成
            sprite_img.putdata(new_data)

            return ImageTk.PhotoImage(sprite_img)
        except Exception as e:
            print(f"スプライト切り出しエラー ({image_name}, index={index}): {e}")
            import traceback
            traceback.print_exc()
            return None

    def on_key_press(self, event):
        """キー押下イベント"""
        mode = self.game.game_state.game_mode

        if mode == GameMode.ENTRANCE:
            self._handle_entrance_keys(event)
        elif mode == GameMode.DUNGEON:
            self._handle_dungeon_keys(event)
        elif mode == GameMode.GAME_OVER or mode == GameMode.GAME_CLEAR:
            if event.keysym == 'x':
                self.game.game_state.game_mode = GameMode.ENTRANCE

    def _handle_entrance_keys(self, event):
        """エントランスのキー処理"""
        if event.keysym == 'Return':
            self.game.start_new_game()
        elif event.keysym == 'z':
            pass  # How to play (未実装)
        elif event.keysym == 'c':
            pass  # Options (未実装)

    def _handle_dungeon_keys(self, event):
        """ダンジョンのキー処理"""
        # 移動キー
        if event.keysym == 'Up':
            self.game.player.direction *= 2
        elif event.keysym == 'Down':
            self.game.player.direction *= 3
        elif event.keysym == 'Right':
            self.game.player.direction *= 5
        elif event.keysym == 'Left':
            self.game.player.direction *= 7

        # 能力キー
        elif event.keysym == 'z':
            msg = self.ability_system.use_all_attack()
            self.add_message(msg)
        elif event.keysym == 'x':
            msg = self.ability_system.use_full_heal()
            self.add_message(msg)
        elif event.keysym == 'c':
            msg = self.ability_system.use_clear_all()
            self.add_message(msg)
        elif event.keysym == 'd':
            msg = self.ability_system.use_monster_removal()
            self.add_message(msg)
        elif event.keysym == 'Return':
            msg = self.ability_system.use_next_floor()
            self.add_message(msg)
        elif event.keysym == 'a':
            msg = self.ability_system.use_open_all_green_boxes()
            self.add_message(msg)

    def add_message(self, message: str, ttl: int = 80):
        """メッセージを追加"""
        if message:
            self.messages.append((message, ttl))

    def update_game(self):
        """ゲームの更新とレンダリング"""
        # プレイヤーの移動処理
        if self.game.game_state.game_mode == GameMode.DUNGEON:
            if self.game.player.direction > 1:
                self.game.move_player(self.game.player.direction)
                self.game.player.direction = 1

                # ダメージメッセージ
                if self.game.game_state.sum_damage > 0:
                    self.add_message(f"プレイヤーは{self.game.game_state.sum_damage}ダメージを受けた")
                    self.game.game_state.sum_damage = 0

        # レンダリング
        self.render()

        # メッセージのTTL減少
        self.messages = [(msg, ttl - 1) for msg, ttl in self.messages if ttl > 0]

        # 次のフレーム
        self.root.after(50, self.update_game)  # 50ms = 20 FPS

    def render(self):
        """画面を描画"""
        self.canvas.delete("all")
        # 画像参照をクリア（前フレームの画像を解放）
        self.canvas._image_refs = []

        mode = self.game.game_state.game_mode

        if mode == GameMode.ENTRANCE:
            self._render_entrance()
        elif mode == GameMode.DUNGEON:
            self._render_dungeon()
        elif mode == GameMode.GAME_OVER:
            self._render_game_over()
        elif mode == GameMode.GAME_CLEAR:
            self._render_game_clear()

    def _render_entrance(self):
        """エントランス画面を描画"""
        self.canvas.create_text(
            self.window_width // 2,
            100,
            text="ダンジョンと不思議の箱",
            fill='white',
            font=self.font,
            justify=tk.CENTER
        )

        y = 200
        self.canvas.create_text(
            self.window_width // 2, y,
            text="ダンジョンに挑む (Enterキー)",
            fill='white',
            font=self.small_font
        )

        y += 50
        self.canvas.create_text(
            self.window_width // 2, y,
            text="← → ↑ ↓ : 移動",
            fill='white',
            font=self.small_font
        )

        y += 30
        self.canvas.create_text(
            self.window_width // 2, y,
            text="Z: 全体攻撃 | X: HP全快 | C: 全消去",
            fill='white',
            font=self.small_font
        )

        y += 30
        self.canvas.create_text(
            self.window_width // 2, y,
            text="D: モンスター除去 | Enter: 次の階",
            fill='white',
            font=self.small_font
        )

        y += 30
        self.canvas.create_text(
            self.window_width // 2, y,
            text="A: 緑箱を全て開ける",
            fill='white',
            font=self.small_font
        )

    def _render_dungeon(self):
        """ダンジョン画面を描画"""
        # マップ領域: 左側 400x400
        map_size = LAND_NUMBER * SC
        floor_variant = (self.game.game_state.floor % 200) // 20  # 壁と部屋のバリエーション

        # タイルを描画
        for i in range(LAND_NUMBER):
            for j in range(LAND_NUMBER):
                square = self.game.map_gen.land_squares[i][j]

                # 画像がある場合は画像を使用
                if self.images:
                    self._draw_tile_sprite(square, floor_variant)
                else:
                    # 画像がない場合は色で描画
                    color = self._get_tile_color(square.condition)
                    self.canvas.create_rectangle(
                        square.left, square.top,
                        square.left + SC, square.top + SC,
                        fill=color,
                        outline=''
                    )

        # プレイヤーを描画
        if self.game.player.alive:
            if self.images:
                self._draw_player_sprite()
            else:
                player_color = self._get_player_color()
                self.canvas.create_oval(
                    self.game.player.left + 2, self.game.player.top + 2,
                    self.game.player.left + SC - 2, self.game.player.top + SC - 2,
                    fill=player_color,
                    outline='white'
                )

        # モンスターを描画
        for monster in self.game.monsters:
            if monster.alive:
                if self.images:
                    self._draw_monster_sprite(monster)
                else:
                    monster_color = self._get_monster_color(monster.ability)
                    self.canvas.create_rectangle(
                        monster.left + 2, monster.top + 2,
                        monster.left + SC - 2, monster.top + SC - 2,
                        fill=monster_color,
                        outline='red'
                    )

        # ステータスバー (下部)
        status_y = map_size + 10
        status_text = (
            f"{self.game.game_state.floor}F "
            f"Lv{self.game.player.level} "
            f"HP {int(self.game.player.hp)}/{self.game.player.max_hp} "
            f"Turn {self.game.game_state.turn}"
        )

        hp_color = 'white'
        if self.game.player.hp <= self.game.player.max_hp / 4:
            hp_color = 'orange'

        self.canvas.create_text(
            10, status_y,
            text=status_text,
            fill=hp_color,
            font=self.small_font,
            anchor=tk.W
        )

        # アビリティバー (右側)
        ability_x = map_size + 10
        ability_y = 10
        line_height = 20

        abilities = [
            f"全攻撃(Z): {self.game.game_state.ability_hp[0]}",
            f"HP全快(X): {self.game.game_state.ability_hp[1]}",
            f"全消去(C): {self.game.game_state.ability_hp[2]}",
            f"モ除去(D): {self.game.game_state.ability_hp[3]}",
            f"次階(Enter): {self.game.game_state.ability_hp[4]}",
            f"復活札: {self.game.game_state.ability_hp[5]}",
        ]

        for i, text in enumerate(abilities):
            self.canvas.create_text(
                ability_x, ability_y + i * line_height,
                text=text,
                fill='white',
                font=self.small_font,
                anchor=tk.W
            )

        # メッセージ表示
        message_y = status_y + 30
        for i, (msg, ttl) in enumerate(self.messages[-3:]):  # 最新3件
            alpha = min(255, ttl * 3)  # フェードアウト効果（簡易版）
            self.canvas.create_text(
                10, message_y + i * 20,
                text=msg,
                fill='yellow',
                font=self.small_font,
                anchor=tk.W
            )

    def _render_game_over(self):
        """ゲームオーバー画面を描画"""
        self.canvas.create_text(
            self.window_width // 2,
            self.window_height // 2 - 50,
            text="GAME OVER",
            fill='red',
            font=self.font,
            justify=tk.CENTER
        )

        self.canvas.create_text(
            self.window_width // 2,
            self.window_height // 2 + 20,
            text="プレイヤーは力尽きた",
            fill='white',
            font=self.small_font
        )

        self.canvas.create_text(
            self.window_width // 2,
            self.window_height // 2 + 60,
            text=f"到達階層: {self.game.game_state.floor}F",
            fill='white',
            font=self.small_font
        )

        self.canvas.create_text(
            self.window_width // 2,
            self.window_height // 2 + 100,
            text="Xキーでメニューに戻る",
            fill='white',
            font=self.small_font
        )

    def _render_game_clear(self):
        """ゲームクリア画面を描画"""
        self.canvas.config(bg='white')

        self.canvas.create_text(
            self.window_width // 2,
            self.window_height // 2 - 50,
            text="GAME CLEAR!",
            fill='blue',
            font=self.font,
            justify=tk.CENTER
        )

        self.canvas.create_text(
            self.window_width // 2,
            self.window_height // 2 + 20,
            text=f"おめでとうございます！{self.game.game_state.floor}階に到達しました！",
            fill='blue',
            font=self.small_font
        )

        self.canvas.create_text(
            self.window_width // 2,
            self.window_height // 2 + 60,
            text=f"HP {int(self.game.player.hp)}/{self.game.player.max_hp} | Lv {self.game.player.level}",
            fill='blue',
            font=self.small_font
        )

        self.canvas.create_text(
            self.window_width // 2,
            self.window_height // 2 + 100,
            text="Xキーでメニューに戻る",
            fill='blue',
            font=self.small_font
        )

        # 背景を元に戻す
        self.root.after(100, lambda: self.canvas.config(bg='black'))

    def _draw_tile_sprite(self, square, floor_variant: int):
        """タイルのスプライトを描画"""
        if square.condition == TileType.WALL:
            # 壁
            sprite = self._get_sprite('wall', floor_variant)
            if sprite:
                self.canvas.create_image(
                    square.left, square.top,
                    image=sprite,
                    anchor=tk.NW
                )
                # 画像の参照を保持（ガベージコレクション防止）
                self.canvas._image_refs = getattr(self.canvas, '_image_refs', [])
                self.canvas._image_refs.append(sprite)

        elif square.condition == TileType.ROOM or square.condition == TileType.ENEMY:
            # 部屋
            sprite = self._get_sprite('room', floor_variant)
            if sprite:
                self.canvas.create_image(
                    square.left, square.top,
                    image=sprite,
                    anchor=tk.NW
                )
                self.canvas._image_refs = getattr(self.canvas, '_image_refs', [])
                self.canvas._image_refs.append(sprite)

        elif square.condition <= TileType.STAIR:
            # 箱または階段
            sprite = self._get_sprite('box', int(square.condition))
            if sprite:
                self.canvas.create_image(
                    square.left, square.top,
                    image=sprite,
                    anchor=tk.NW
                )
                self.canvas._image_refs = getattr(self.canvas, '_image_refs', [])
                self.canvas._image_refs.append(sprite)

    def _draw_player_sprite(self):
        """プレイヤーのスプライトを描画"""
        # プレイヤーの状態に応じたスプライトインデックス
        sprite_index = int(self.game.player.condition)
        sprite = self._get_sprite('player', sprite_index)

        if sprite:
            self.canvas.create_image(
                self.game.player.left, self.game.player.top,
                image=sprite,
                anchor=tk.NW
            )
            self.canvas._image_refs = getattr(self.canvas, '_image_refs', [])
            self.canvas._image_refs.append(sprite)

    def _draw_monster_sprite(self, monster):
        """モンスターのスプライトを描画"""
        # モンスターのアビリティに応じたスプライトインデックス
        sprite_index = int(monster.ability)
        sprite = self._get_sprite('monster', sprite_index)

        if sprite:
            self.canvas.create_image(
                monster.left, monster.top,
                image=sprite,
                anchor=tk.NW
            )
            self.canvas._image_refs = getattr(self.canvas, '_image_refs', [])
            self.canvas._image_refs.append(sprite)

    def _get_tile_color(self, tile_type: TileType) -> str:
        """タイルの色を取得"""
        color_map = {
            TileType.BLUE_BOX: 'blue',
            TileType.RED_BOX: 'red',
            TileType.YELLOW_BOX: 'yellow',
            TileType.GREEN_BOX: 'green',
            TileType.PURPLE_BOX: 'purple',
            TileType.STAIR: 'cyan',
            TileType.WALL: 'gray',
            TileType.ROOM: 'black',
            TileType.ENEMY: 'black',  # モンスターは別途描画
        }
        return color_map.get(tile_type, 'black')

    def _get_player_color(self) -> str:
        """プレイヤーの色を状態で変更"""
        if self.game.player.condition == PlayerCondition.STEALTH:
            return 'lightgray'  # 透明（薄い色）
        elif self.game.player.condition == PlayerCondition.WALL_BREAK:
            return 'orange'
        elif self.game.player.condition == PlayerCondition.SLOW:
            return 'lightblue'
        else:
            return 'white'

    def _get_monster_color(self, ability: MonsterAbility) -> str:
        """モンスターの色をアビリティで変更"""
        color_map = {
            MonsterAbility.NO_ABILITY: 'darkred',
            MonsterAbility.WALL_BREAK: 'orange',
            MonsterAbility.SLOW: 'lightblue',
            MonsterAbility.BOX_ATTACK: 'brown',
            MonsterAbility.STEALTH: 'lightgray',
        }
        return color_map.get(ability, 'darkred')


def main():
    """メイン関数"""
    root = tk.Tk()
    app = DungeonGameGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
