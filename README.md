# ダンジョンと不思議の箱 - Python版

Visual Basicで開発されたオリジナルゲーム「ダンジョンと不思議の箱」のPython版リファクタリング実装です。

## 概要

このプロジェクトは、元のVisual Basicコードの機能を維持しながら、Pythonの現代的なプログラミング手法を使用してリファクタリングしたものです。

### 特徴

- **ターン制ダンジョン探索RPG**
- **40x40のランダム生成マップ**
- **様々な効果を持つ不思議な箱システム**
- **モンスターとの戦闘**
- **特殊能力システム**
- **1000階を目指すエンドレスダンジョン**

## ファイル構成

### 元のVisual Basicファイル
- `Project1.vbp` - プロジェクトファイル
- `Form1.frm` - メインフォーム
- `BasicModule.bas` - 基本モジュール
- `LandModule.bas` - 地形管理モジュール
- `MonsterModule.bas` - モンスター管理モジュール
- `AbilityModule.bas` - 特殊能力モジュール
- `WordModule.bas` - テキスト管理モジュール
- `ImscLib12.cls` - サウンドライブラリクラス

### Python実装ファイル
- `dungeon_game.py` - **メインゲームエンジン**
- `console_ui.py` - **コンソールUI実装**
- `README.md` - このファイル

### リソースファイル
- `Images/` - ゲーム画像（BMP形式）
- `Sounds/` - ゲーム音楽（MP3形式）

## 実行方法

### 必要要件
- Python 3.7以上

### コアエンジンの実行
```bash
python3 dungeon_game.py
```

### コンソール版ゲームのプレイ
```bash
python3 console_ui.py
```

## ゲームシステム

### ゲームモード
1. **Entrance** - タイトル画面
2. **Dungeon** - メインゲーム
3. **HowtoPlay** - 操作説明
4. **Options** - 設定画面
5. **GameOver** - ゲームオーバー
6. **GameClear** - ゲームクリア（1000階到達）
7. **Museum** - 記録閲覧

### 操作方法
- **移動**: 矢印キー (↑↓←→)
- **特殊能力**:
  - `Z`: 全体攻撃
  - `X`: HP全回復
  - `C`: 全体回復
  - `D`: 全モンスター除去
  - `Enter`: 次の階層へ
  - `A`: 緑箱効果発動

### 箱の種類と効果

#### 青箱 (BlueBox)
- HP小回復
- モンスターの状態異常解除

#### 赤箱 (RedBox)
- 悪い効果（15種類のランダム効果）
- HP1にする、攻撃力低下、状態異常など

#### 黄箱 (YellowBox)
- 様々な効果（15種類のランダム効果）
- 良い効果と悪い効果が混在

#### 緑箱 (GreenBox)
- 良い効果（15種類のランダム効果）
- HP全回復、能力向上、モンスター弱体化など

#### 紫箱 (PurpleBox)
- 特殊能力の使用回数を増加
- 難易度に応じて増加量が変動

### プレイヤー能力

#### 基本ステータス
- **HP**: ヒットポイント
- **ATK**: 攻撃力
- **DEF**: 防御力
- **Level**: レベル（経験値で上昇）
- **Exp**: 経験値

#### 特殊状態
- **WallBreak**: 壁破壊能力
- **Slow**: 移動速度低下
- **Stealth**: 透明状態（モンスターに発見されない）

### モンスターシステム

#### モンスター能力
- **NoAbility**: 通常
- **WallBreak**: 壁破壊
- **Slow**: 移動速度低下
- **BoxAttack**: 箱攻撃
- **Stealth**: 透明状態

#### AI行動
- プレイヤーに向かって移動
- 特殊能力に応じた行動
- ターン制での移動

## 技術仕様

### アーキテクチャ

#### クラス設計
```python
# 基本クラス
Status          # 基本状態クラス
Position        # 位置情報
Rectangle       # 矩形領域

# ゲームオブジェクト
Player          # プレイヤークラス
Monster         # モンスタークラス
LandSquare      # 地形スクエアクラス

# ゲーム管理
GameState       # ゲーム状態管理
GameEngine      # ゲームエンジン

# UI
ConsoleUI       # コンソールUI
```

#### 列挙型 (Enum)
```python
GameMode        # ゲームモード
LandCondition   # 地形条件
Ability         # 特殊能力
Difficulty      # 難易度
```

### データ構造

#### ゲーム状態
- 40x40の2次元配列によるマップ管理
- 最大1000体のモンスター管理
- プレイヤー状態と各種パラメータ
- UI表示用テキスト配列

#### 戦闘システム
- ダメージ計算: `ATK * (0.9~1.1のランダム) / DEF`
- 経験値によるレベルアップ: `必要経験値 = Level^3`
- ステータス上限管理

## 元のVBコードとの対応

### モジュール対応表
| VBファイル | Python対応クラス/関数 |
|-----------|---------------------|
| BasicModule.bas | GameEngine, Status |
| LandModule.bas | LandSquare, GameState._land_set |
| MonsterModule.bas | Monster, GameState._monster_set |
| AbilityModule.bas | GameEngine._ability_effect |
| WordModule.bas | GameState._word_set |
| Form1.frm | ConsoleUI (UI部分) |

### 主要関数の対応
| VB関数 | Python対応メソッド |
|--------|------------------|
| Movement() | GameEngine._movement() |
| StatusCheck() | GameEngine._status_check() |
| FloorSet() | GameEngine._floor_set() |
| BattleCheck() | GameEngine._battle_check_* |
| BoxEffect() | GameEngine._box_effect() |
| PositionCheck() | GameEngine._position_check_* |

## 拡張可能性

### 今後の拡張予定
1. **グラフィカルUI** - pygame/tkinterを使用したGUI
2. **サウンド対応** - 元のMP3ファイルの再生
3. **セーブ/ロード機能** - ゲーム進行状況の保存
4. **難易度設定** - より詳細な難易度調整
5. **追加コンテンツ** - 新しい箱や能力の追加

### GUI実装例（pygame使用）
```python
import pygame

class PygameUI:
    def __init__(self):
        pygame.init()
        self.screen = pygame.display.set_mode((800, 600))
        # ... 実装
```

## ライセンス

このプロジェクトは教育目的で作成されました。
元のVisual Basicゲームの作者に敬意を表します。

音楽データは「フリー音楽 H/MIX GALLERY」(http://www.hmix.net/) から提供されています。

## 開発者向け情報

### コードスタイル
- PEP 8準拠
- 型ヒント使用
- dataclassesとenumを活用
- 適切なコメント記述

### テスト
```bash
# 基本動作テスト
python3 dungeon_game.py

# インタラクティブテスト
python3 console_ui.py
```

### デバッグ機能
- ゲーム状態の可視化
- コンソール出力でのデバッグ情報
- 各種パラメータの表示

このPython版実装により、元のVisual Basicゲームの魅力を現代的な環境で楽しむことができます。
