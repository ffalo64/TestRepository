#!/usr/bin/env python3
"""
ダンジョンと不思議の箱 - コンソールUI実装

コンソールベースのユーザーインターフェースでゲームをプレイできます。
"""

import os
import sys
from dungeon_game import GameEngine, GameMode, LandCondition, Ability


class ConsoleUI:
    """コンソールUIクラス"""

    # 表示用記号
    SYMBOLS = {
        LandCondition.BLUE_BOX: '🔵',    # 青箱
        LandCondition.RED_BOX: '🔴',     # 赤箱
        LandCondition.YELLOW_BOX: '🟡',  # 黄箱
        LandCondition.GREEN_BOX: '🟢',   # 緑箱
        LandCondition.PURPLE_BOX: '🟣', # 紫箱
        LandCondition.STAIR: '⬇️',       # 階段
        LandCondition.WALL: '🟫',        # 壁
        LandCondition.ROOM: '  ',        # 部屋
        LandCondition.ENEMY: '👾',       # 敵
    }

    # ASCIIフォールバック記号（絵文字が表示できない環境用）
    ASCII_SYMBOLS = {
        LandCondition.BLUE_BOX: 'B ',
        LandCondition.RED_BOX: 'R ',
        LandCondition.YELLOW_BOX: 'Y ',
        LandCondition.GREEN_BOX: 'G ',
        LandCondition.PURPLE_BOX: 'P ',
        LandCondition.STAIR: 'S ',
        LandCondition.WALL: '##',
        LandCondition.ROOM: '. ',
        LandCondition.ENEMY: 'M ',
    }

    def __init__(self, use_ascii=False):
        """
        初期化

        Args:
            use_ascii: Trueの場合、絵文字の代わりにASCII文字を使用
        """
        self.engine = GameEngine()
        self.use_ascii = use_ascii
        self.symbols = self.ASCII_SYMBOLS if use_ascii else self.SYMBOLS
        self.running = True

    def clear_screen(self):
        """画面クリア"""
        os.system('clear' if os.name == 'posix' else 'cls')

    def render_entrance(self):
        """エントランス画面の描画"""
        self.clear_screen()
        print("=" * 60)
        print(" " * 15 + "ダンジョンと不思議の箱")
        print("=" * 60)
        print()
        print("  📜 ダンジョンに入る (Enter)")
        print("  ❓ 遊び方を見る (H)")
        print("  ⚙️  オプション (O)")
        print("  🚪 終了 (Q)")
        print()
        print("=" * 60)
        print()
        print("音楽データ提供元:")
        print("  フリー音楽 H/MIX GALLERY")
        print("  管理者: H♪M♭")
        print("  http://www.hmix.net/")
        print()

    def render_dungeon(self):
        """ダンジョン画面の描画"""
        self.clear_screen()
        state = self.engine.state

        # タイトル
        print("=" * 80)
        print(f" フロア: {state.floor}F | " +
              f"Lv.{state.player.level} | " +
              f"HP: {int(state.player.hp)}/{state.player.max_hp} | " +
              f"Turn: {state.turn}")
        print("=" * 80)

        # マップ表示（簡易版：プレイヤー周辺のみ表示）
        self._render_map_around_player()

        # ステータス表示
        print("=" * 80)
        self._render_status()

        # メッセージ表示
        print("-" * 80)
        self._render_messages()
        print("-" * 80)

        # 操作説明
        print()
        print("移動: ↑↓←→ (WASD) | 全体攻撃(Z) | HP回復(X) | 全除去(C) | モンスター除去(D)")
        print("次の階(Enter) | 緑箱発動(A) | 終了(Q)")

    def _render_map_around_player(self):
        """プレイヤー周辺のマップを描画"""
        state = self.engine.state
        sc = state.SC
        ln = state.LAND_NUMBER

        # プレイヤーの位置
        px, py = state.player.left // sc, state.player.top // sc

        # 表示範囲（プレイヤー中心に15x15）
        view_range = 7
        start_x = max(0, px - view_range)
        end_x = min(ln, px + view_range + 1)
        start_y = max(0, py - view_range)
        end_y = min(ln, py + view_range + 1)

        print()
        for y in range(start_y, end_y):
            line = ""
            for x in range(start_x, end_x):
                # プレイヤー位置
                if x == px and y == py:
                    line += "😀" if not self.use_ascii else "P "
                else:
                    # モンスターチェック
                    is_monster = False
                    for m in state.monsters:
                        if m.alive and m.left // sc == x and m.top // sc == y:
                            line += self.symbols[LandCondition.ENEMY]
                            is_monster = True
                            break

                    if not is_monster:
                        # 地形
                        land = state.land_squares[x][y]
                        symbol = self.symbols.get(land.condition, '? ')
                        line += symbol

            print("  " + line)
        print()

    def _render_status(self):
        """ステータス情報の表示"""
        state = self.engine.state
        p = state.player

        print(f"【プレイヤー】")
        print(f"  HP: {int(p.hp)}/{p.max_hp} | " +
              f"ATK: {int(p.atk)} | DEF: {int(p.defense)} | " +
              f"EXP: {p.exp}/{p.level ** 3}")

        print()
        print(f"【特殊能力】")
        print(f"  全体攻撃(Z): {state.ability_hp[0]}回 | " +
              f"HP全回復(X): {state.ability_hp[1]}回 | " +
              f"全除去(C): {state.ability_hp[2]}回")
        print(f"  モンスター除去(D): {state.ability_hp[3]}回 | " +
              f"次階(Enter): {state.ability_hp[4]}回 | " +
              f"復活の像: {state.ability_hp[5]}回")

    def _render_messages(self):
        """メッセージの表示"""
        state = self.engine.state

        print("【メッセージ】")
        for i in range(4):
            if state.words[i]:
                print(f"  {state.words[i]}")

    def render_game_over(self):
        """ゲームオーバー画面"""
        self.clear_screen()
        print()
        print("=" * 60)
        print(" " * 20 + "GAME OVER")
        print("=" * 60)
        print()
        print(f"  到達フロア: {self.engine.state.floor}F")
        print(f"  最終レベル: Lv.{self.engine.state.player.level}")
        print()
        print("  タイトルに戻る (Enter)")
        print("  終了 (Q)")
        print()
        print("=" * 60)

    def render_game_clear(self):
        """ゲームクリア画面"""
        self.clear_screen()
        print()
        print("=" * 60)
        print(" " * 15 + "🎉 GAME CLEAR 🎉")
        print("=" * 60)
        print()
        print(f"  おめでとうございます！1000階に到達しました！")
        print()
        print(f"  最終レベル: Lv.{self.engine.state.player.level}")
        print(f"  最終HP: {int(self.engine.state.player.hp)}/{self.engine.state.player.max_hp}")
        print(f"  最終ATK: {int(self.engine.state.player.atk)}")
        print(f"  最終DEF: {int(self.engine.state.player.defense)}")
        print()
        print("  タイトルに戻る (Enter)")
        print("  終了 (Q)")
        print()
        print("=" * 60)

    def render_help(self):
        """ヘルプ画面"""
        self.clear_screen()
        print("=" * 60)
        print(" " * 20 + "遊び方")
        print("=" * 60)
        print()
        print("【ゲーム概要】")
        print("  ダンジョンの1000階を目指すローグライクRPGです。")
        print("  モンスターを倒しながら階段を探して進みましょう！")
        print()
        print("【操作方法】")
        print("  ↑↓←→ または WASD : 移動")
        print("  Z : 全体攻撃（全モンスターにダメージ）")
        print("  X : HP全回復")
        print("  C : 全除去（箱・壁・モンスターを全て除去）")
        print("  D : モンスター除去（全モンスターを箱に変換）")
        print("  Enter : 次の階へ強制移動")
        print("  A : 全ての緑箱の効果を発動")
        print()
        print("【箱の種類】")
        print("  🔵 青箱   : HP小回復 + モンスター状態異常解除")
        print("  🔴 赤箱   : 悪い効果（15種類）")
        print("  🟡 黄箱   : 良い効果と悪い効果が混在（15種類）")
        print("  🟢 緑箱   : 良い効果（15種類）")
        print("  🟣 紫箱   : コマンド使用回数増加")
        print()
        print("【ヒント】")
        print("  - ターンが0になるとゲームオーバーです")
        print("  - 箱を壊すとターンが10増えます")
        print("  - レベルアップで能力が上昇します")
        print()
        print("タイトルに戻る (Enter)")

    def handle_input_entrance(self, key: str):
        """エントランス画面の入力処理"""
        if key.lower() == 'q':
            self.running = False
        elif key == '\n' or key == '\r':  # Enter
            self.engine.handle_key('enter')
        elif key.lower() == 'h':
            self.render_help()
            self._wait_for_enter()

    def handle_input_dungeon(self, key: str):
        """ダンジョン画面の入力処理"""
        key_map = {
            'w': 'up',
            's': 'down',
            'a': 'left',
            'd': 'right',
            '\x1b[A': 'up',     # 上矢印
            '\x1b[B': 'down',   # 下矢印
            '\x1b[C': 'right',  # 右矢印
            '\x1b[D': 'left',   # 左矢印
            'z': 'z',
            'x': 'x',
            'c': 'c',
            '\n': 'enter',
            '\r': 'enter',
        }

        if key.lower() == 'q':
            self.running = False
            return

        # キーマッピング
        game_key = key_map.get(key.lower())
        if game_key:
            self.engine.handle_key(game_key)
            self.engine.update()

    def handle_input_game_over(self, key: str):
        """ゲームオーバー画面の入力処理"""
        if key.lower() == 'q':
            self.running = False
        elif key == '\n' or key == '\r':
            self.engine.state.reset_entrance()

    def _wait_for_enter(self):
        """Enterキー待機"""
        try:
            import termios
            import tty
            fd = sys.stdin.fileno()
            old_settings = termios.tcgetattr(fd)
            try:
                while True:
                    tty.setraw(fd)
                    ch = sys.stdin.read(1)
                    if ch == '\n' or ch == '\r':
                        break
            finally:
                termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
        except:
            input()

    def get_key(self):
        """キー入力を取得（非ブロッキング）"""
        try:
            import termios
            import tty
            fd = sys.stdin.fileno()
            old_settings = termios.tcgetattr(fd)
            try:
                tty.setraw(fd)
                ch = sys.stdin.read(1)

                # エスケープシーケンス処理（矢印キー）
                if ch == '\x1b':
                    ch += sys.stdin.read(2)

                return ch
            finally:
                termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
        except:
            # Windows環境など、termiosが使えない場合
            return input()

    def run(self):
        """メインループ"""
        print("ダンジョンと不思議の箱 - Python版")
        print("ロード中...")

        import time
        time.sleep(0.5)

        while self.running:
            # 画面描画
            if self.engine.state.game_mode == GameMode.ENTRANCE:
                self.render_entrance()
            elif self.engine.state.game_mode == GameMode.DUNGEON:
                self.render_dungeon()
            elif self.engine.state.game_mode == GameMode.GAME_OVER:
                self.render_game_over()
            elif self.engine.state.game_mode == GameMode.GAME_CLEAR:
                self.render_game_clear()

            # 入力処理
            try:
                key = self.get_key()
            except KeyboardInterrupt:
                self.running = False
                break

            # モード別入力処理
            if self.engine.state.game_mode == GameMode.ENTRANCE:
                self.handle_input_entrance(key)
            elif self.engine.state.game_mode == GameMode.DUNGEON:
                self.handle_input_dungeon(key)
            elif self.engine.state.game_mode == GameMode.GAME_OVER:
                self.handle_input_game_over(key)
            elif self.engine.state.game_mode == GameMode.GAME_CLEAR:
                self.handle_input_game_over(key)

        # 終了処理
        self.clear_screen()
        print()
        print("=" * 60)
        print(" " * 15 + "ゲームを終了します")
        print(" " * 10 + "遊んでいただきありがとうございました！")
        print("=" * 60)
        print()


def main():
    """メイン関数"""
    import sys

    # コマンドライン引数で ASCII モード選択
    use_ascii = '--ascii' in sys.argv or '-a' in sys.argv

    if use_ascii:
        print("ASCII モードで起動します")
    else:
        print("絵文字モードで起動します")
        print("絵文字が正しく表示されない場合は --ascii オプションを使用してください")

    ui = ConsoleUI(use_ascii=use_ascii)

    try:
        ui.run()
    except Exception as e:
        print(f"\nエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
        print("\nゲームを終了します")


if __name__ == "__main__":
    main()
