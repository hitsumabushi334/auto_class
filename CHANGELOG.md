# Changelog

## [Unreleased]

### Added

- `config.json`: `app` セクションを追加（`note_output_mode`: "MD" or "Word"、`developer_mode`: true/false）。
- `config_manager.py`: `update_config(new_config)` メソッド追加 — インメモリ状態とファイルを同時更新（変更前の古い値が残る問題を解消）。
- `config_manager.py`: `reset_to_default()` メソッド追加 — `DEFAULT_CONFIG` で上書きし即時保存。
- `config_manager.py`: `get_current_config()` メソッド追加 — ディープコピーを返す。
- `config_manager.py`: `get_app_settings()` メソッド追加 — `note_output_mode`/`developer_mode` を返す。
- `main.py`: `ttk.Notebook` によるタブ化（「Home」「設定」2タブ）。既存UI要素は Home タブに移動。
- `main.py`: 設定タブ — Canvas+Scrollbar によるスクロール可能フォーム。全パラメータ（API/音声/スクショ/UI/ログ/アプリ）を編集可能にし、各パラメータに説明文を付与。
- `main.py`: 設定保存処理 — バリデーション → `ConfigManager.update_config()` → アプリ内変数の即時反映（閾値・APIクライアント再初期化・ウィンドウサイズ適用・ロガーハンドラ切り替え）。
- `main.py`: リセット処理 — `ConfigManager.reset_to_default()` → UI・変数の即時反映。
- `main.py`: タブロック機能 — `start_tasks` で設定タブを無効化し、ノート生成完了（`_restore_gui_after_note_creation`）まで設定変更を防止。
- `main.py`: ノート出力モード分岐 — `note_output_mode == "Word"` 時に `.docx` を生成（旧コメントアウトロジックを復活・`term.get("md","")` を `term.get("word","")` に修正）、それ以外は `.md` を生成。
- `main.py`: 開発者モード制御 — 起動時・保存時に `StreamHandler` を動的追加/削除。
- `tests/test_config_manager.py`: TDD による新機能テスト13ケース追加（合計25ケース全パス）。
- `main.py` [Fix]: 設定タブで `api.models` を保存時、有効な辞書リスト形式かバリデーションするよう修正 (Codex P1指摘)。
- `main.py` [Fix]: 録画開始処理 (`start_tasks`) 失敗時に、ロックした設定タブを即座に解除・復元する処理を追加 (Codex P2指摘)。
- `main.py` [Fix]: 設定保存時、Gemini APIキーが無効・空ならAPIクライアントの初期化をクリアし、エラーを伝搬させないよう修正 (Codex P2指摘)。
- `main.py` [Fix]: 設定保存時、Geminiモデル一覧の更新に伴いコンボボックスの選択値を追従し、リスト外になればデフォルト値に戻す処理を追加 (Codex P2指摘)。
- `main.py` [Fix]: 設定保存時にログレベル、フォーマット、FileHandlerとStreamHandlerの動的再設定を完全に行い、即時反映されるよう修正 (Codex P3指摘)。

### Changed

- `config_manager.py`: `api.models` を文字列リストからオブジェクトリスト（`[{"model_name": "...", "description": "..."}]`）へ変更。`model_descriptions` 辞書を廃止。
- `config_manager.py`: 設定ファイルが存在しない・パース失敗・旧フォーマット検出時に、デフォルト設定 (`DEFAULT_CONFIG`) で自動生成/上書きし、エラーログを出力するよう改善。PermissionError 発生時はインメモリのデフォルト設定で起動を継続。
- `config_manager.py`: `get_api_key()` を環境変数 (`GEMINI_API_KEY`) への依存を廃止し、`config.json` からの読み込みに一本化。
- `main.py`: `from dotenv import load_dotenv` インポートおよび `load_dotenv()` 呼び出しを削除。
- `CONFIG.md`: 新しい models 構造・APIキー設定方法・自動生成機能の説明に更新。
- `.gitignore`: `config.json` を除外対象に追加（APIキー流出防止）。

### Added

- `config_manager.py`: モジュールレベルの `DEFAULT_CONFIG` 辞書を追加（テストや初期生成のソースとして使用）。
- `config.example.json`: APIキー設定のテンプレートファイルとして追加。
- `tests/test_config_manager.py`: TDD で作成したユニットテスト (11ケース)。生成・上書き・権限エラー・ゲッターの動作を網羅。
- `docs/config-json-update/implementation_plan.md`: 実装計画書を追加。
