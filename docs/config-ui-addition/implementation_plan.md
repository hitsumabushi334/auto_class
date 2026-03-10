# Implementation Plan: Config UI Addition

## Overview

`config.json` に `note_output_mode` (word or MD) と `developer_mode` (コンソール出力有無) を追加し、メインのUIをタブ化（Homeタブ・設定タブ）に分割します。設定タブから各種パラメータの設定変更、保存、デフォルトへのリセットを行えるようにします。

## Background / Context

ユーザーの追加要望により、UIの整理と設定変更の利便性向上が求められています。特にノート生成でWordかMarkdownかを選べるようにし、さらに開発者向けのコンソール出力切り替えができるようにします。現状全て1画面に表示されているUIをタブ分けし、スッキリさせます。

## Goals

1. `config.json` への新規パラメータ（`app.note_output_mode`, `app.developer_mode`）の追加
2. `main.py` のUIを `ttk.Notebook` を用いて「Home」タブと「設定」タブに分割（既存要素はHomeタブへ移動）
3. **設定タブにて、`config.json` に存在する「すべてのパラメータ」を編集可能にする**（`api.models` や `ui`, `logging` 関連パラメータなども含む）
4. **設定変更の保存時、即時に `main.py` 内の変数に反映させる**。また保存先は `ConfigManager` を経由した状態保存ではなく、**直接 `config.json` ファイルを上書き保存**（`json.dump`）する構造とする
5. デフォルト設定へのリセット機能の実装（リセット時も即時反映・ファイル保存）
6. **録画中・処理中（ノート作成など）は、タブの切り替え自体を無効化（ブロック）する**
7. ノート生成処理にて `app.note_output_mode` を参照し、`Word` (docx) と `MD` の出力を分岐させる（Word出力はコメントアウトされていたロジックを流用する）
8. `app.developer_mode` がオンの場合のみコンソールログ出力（StreamHandler等）を有効化し、オフの場合はファイル出力のみにする

## Non-Goals

- note_templete と現在のノート作成ロジックの差異修正（後で手直し予定と伺っているため今回は現状のコメントアウトの復旧に留める）
- アプリケーション全体のGUIデザインの根本的な変更

## Design / Approach

- `config_manager.py`:
  - `DEFAULT_CONFIG` 配下の `app` セクションに `note_output_mode` (デフォルト "MD") と `developer_mode` (デフォルト False) を追加。
  - 新たに `update_config(self, new_config_dict)` および `save_config(self)` を実装する。
  - **更新処理**: `update_config` を呼ばれた際に、自身の保持する `self._config` を新しい値で上書きし、即座に `config.json` ファイルへ書き込む（ご指摘の「変更前の古い値が保持されてしまう問題」は、ここで `ConfigManager` 内のインメモリ変数 `self._config` も合わせて更新することで解決します）。
  - リセット用の `reset_to_default(self)` も実装し、同様に `self._config` を `DEFAULT_CONFIG` で上書きしてファイル保存する。
- `main.py`:
  - **変数の保持と即時反映**:
    - `__init__` で `ConfigManager` から設定を読み込み、インスタンス変数（`self.similarity_threshold` やGeminiのモデル名など）に保持する。
  - **UI改修: タブ化**:
    - 既存のUIを `ttk.Notebook` の `home_tab` (Frame) にパッキング。
    - `settings_tab` (Frame) を作成し、CanvasとScrollbarを組み合わせてスクロール可能な設定領域を構築。
  - **設定タブ実装**:
    - `api.gemini_api_key`, `api.models` (JSONテキストエリア等で編集), `api.default_model_index`, `audio.*`, `screenshot.*`, `ui.*`, `logging.*`, `app.*` の全項目を入力ウィジェット（Entry, Combobox, Textなど）として配置。
    - 各パラメータの簡易説明（「大きくすると類似度の判定が厳しくなります」等）をLabelとして併記。
  - **保存ロジック**:
    - 「保存」ボタン押下時、全ウィジェットから値を取得しバリデーションを実施。
    - 通過した場合、**取得した全データを辞書化し、`get_config_manager().update_config(new_dict)` に渡す（ファイル保存と ConfigManager の状態更新が同時に行われる）**。
    - 直後に、**`main.py` 自身の各状態変数へ新しい値を代入し、即時反映させる**（Geminiのモデル再初期化、各種閾値の更新、ウィンドウサイズの `root.geometry` 再適用、ロギングハンドラの再設定など）。これにより、アプリを再起動せずとも次回の録画・キャプチャから新しい設定値で動作する。
  - **タブの切り替え制限**:
    - 録画開始時、および後続の処理（ノート生成など）の間、`self.notebook.tab(1, state="disabled")` を呼び出し設定タブを開けなくする。処理が完全に完了した時点で `state="normal"` に戻す。
  - **ノート生成処理（MD / Word出力の切り替え）**:
    - `app.note_output_mode` （`self.config_manager.get(...)` ではなくキャッシュされた即時反映済みの変数）を確認し、`Word` の場合は `from docx import Document` およびコメントアウトされていたロジックを有効化し保存。`term.get("md", "")` は `term.get("word", "")` に修正。
  - **開発者モード用ログ制御処理**:
    - `developer_mode` の値をもとに、GUI起動時や設定保存時に `logging.getLogger().handlers` を走査し、False なら StreamHandler を削除（またはレベル変更）、True なら追加（存在しなければ）するよう動的に制御する。

## Implementation Steps

1. `config_manager.py` の改修
   - `DEFAULT_CONFIG` へ `app` 項目追加
   - 更新処理 `update_config` (内部変数更新 + JSON上書き保存)、および `reset_to_default` の実装
2. `main.py` の変数保持の整理と UI タブ化
3. `main.py` の設定タブUI（全パラメータのフォーム、Canvasスクロールエリア、簡易説明文）の実装
4. `main.py` : 設定保存処理（バリデーション -> `ConfigManager` 更新 -> `main.py` 各変数の即時同期）の実装
5. `main.py` : リセット処理の実装
6. タブ無効化処理の追加（録画・後処理の開始から完了までロック）
7. ノート生成処理の分岐追加（MD / Word出力）と コンソールログの表示制御（ハンドラの動的追加・削除）

## Testing Strategy

- **手動確認**:
  1. タブが2つ（Home, 設定）存在すること。
  2. 設定タブで存在する全パラメータが編集可能であること。
  3. 正常な値を設定して「保存」した際、`config.json` が上書きされること。
  4. 保存直後にサイズ変更や閾値などの設定が即座にメインプロセスに効いていること（アプリ再起動なしで）。
  5. 録画を開始した状態から、ノート生成処理が完了するまでの一連のフロー全体を通して設定タブへ移動できなくなっていること。
  6. ノート出力をWordに変更し、処理を実行して `.docx` ファイルが出力されるか確認。MDのときは `.md` が出力されるか確認。
  7. 開発者モード変更時にコンソールログ出力が動的に切り替わる（ONにするとターミナルに出力され、OFFにすると出力されない）こと。

## Risks & Mitigations

- `ConfigManager` を経由せず `main.py` で直接保存することによる責務分散と保守性低下のリスク（Codex指摘事項）。**Mitigation**: 今回はユーザーの強い要望によりこの構成をとるが、設定の読み書き関数を `main.py` 内でなるべく一箇所にまとめ、散逸を防ぐ。
- `developer_mode` の切り替え時にロガーハンドラが重複したり消えたりする問題（Codex指摘事項）。**Mitigation**: 設定適用時に現在の `handlers` をチェックし、重複登録を防ぐと共に、安全に再設定するロジックを組む。
- `docx` （python-docx） を使用した保存時、現在コメントアウトされているコード `term.get("md", "")` に問題がある。**Mitigation**: 実装時に `term.get("word", "")` に合わせて微調整を行う。

## Open Questions

- 特になし。
