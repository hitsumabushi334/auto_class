# Changelog

## [Unreleased]

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
