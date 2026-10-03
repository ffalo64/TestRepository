# ダンジョンと不思議の箱 — VB6 → ブラウザ移植

## 構成
- ルート直下の `*.bas` / `Form1.frm` / `ImscLib12.cls` が元の VB6 ソース（Shift_JIS）。**参照専用・編集しない**。
- `Images/` `Sounds/` が元アセット。ブラウザ版は `web-game/public/` にコピーしたものを使う。
- `web-game/` がブラウザ版（Vite + 素の ES modules）。
  - `npm run dev` → http://localhost:5173/
  - `npm test` → Node でのヘッドレス動作テスト（`test/headless.js`）
  - `npm run autoplay [games] [maxTicks]` → オートプレイ AI の成績（到達階・死因・100F 到達率）を集計（`test/autoplay.js`。環境変数 `GOAL` `STOP` `SEED0`）
  - `src/engine.js` は DOM 非依存のゲームロジック（VB6 の各 Sub を移植）。描画は `renderer.js`、音は `audio.js`、入力は `main.js`。
  - `src/autoplay.js` はブラウザ版独自のオートプレイ AI（P キーで OFF → 等速 → 4 倍速）。DOM 非依存で、`nextKeys(g)` が返すキーを `keyDown()` に渡す。エンジンには手を入れない。
    戦略は箱の期待値（ターン換算）と Dijkstra（箱を壊して進む経路込み）で目的地を選び、戦術は近くのモンスターの動きを 4 手先までシミュレーションして被弾を避ける。100F 到達率は約 40%（48 ゲーム）。

## 移植方針
- ゲームロジックは VB6 に忠実に。VB6 の Double→Long 代入は銀行丸め（`cLng`）、`Int()` は `Math.floor`。
- 意図的に変えた点はコード中に `// port:` コメントで明記する（範囲外参照のクラッシュ回避など）。

## 作業の振り分け（モデル選択）
メインエージェント（Opus）は設計・ゲームロジック移植・レビュー・難しいデバッグに集中し、**単純な作業はサブエージェントに回す**。
- **Haiku**（`model: "haiku"`）: ファイルのコピー/移動、雛形作成（package.json 等）、`npm install`、ビルド/テスト実行と結果報告、grep による単純な調査、定型的な置換。
- **Sonnet**（`model: "sonnet"`）: 仕様が明確な小〜中規模の実装（テストハーネス、UI の小修正、ドキュメント作成）、単純なバグ修正。
- **Opus（メイン）**: VB6 ロジックの解釈・移植、アーキテクチャ判断、サブエージェント成果物の確認。
- サブエージェントには目的・対象パス・完了条件を具体的に書いて渡し、戻ってきた成果は必ずメインで確認する。
