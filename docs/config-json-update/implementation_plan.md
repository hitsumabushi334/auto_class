# 実装計画書: config.json の構造変更と自動生成機能の追加

## 目的

1. **設定ファイル構造の更新**: `json.api.models` の構造を文字列のリストからオブジェクトのリスト（`[{"model_name": "...", "description": "..."}]`）へ変更し、独立していた `model_descriptions` 辞書を廃止する。
2. **自動生成機能の追加**: アプリケーション起動時に `config.json` が存在しない、あるいは読み込みエラー（旧形式、パース失敗等）が発生した場合、デフォルト値を含む `config.json` で上書き・生成してから起動する仕様に変更する。上書き時は明確なエラーログを出力する。
3. **APIキー取得元の単一化と保護**:
   - APIキーの取得処理において、環境変数の使用を完全に廃止し、設定ファイル（`config.json`）の `api.gemini_api_key` からの読み込みへと一本化する。
   - `config.json` を `.gitignore` に追加してコミット対象外とし、代わりに `config.example.json` をリポジトリに含めることでAPIキーの流出を防ぐ。
4. **既存処理の修正**: 仕様変更に基づき、`config_manager.py` だけでなく `main.py` においても不要な環境変数ロード処理を削除するなどの修正を行う。

## 提案する変更内容

### 設定管理

#### [MODIFY] `config_manager.py`(file:///c:/auto_class/config_manager.py)

- **デフォルト設定の定義**: クラスまたはモジュールレベルで `DEFAULT_CONFIG` 辞書を新構造に基づいて定義する。
- **`load_config` 関数の修正**:
  - `config_path` が存在しない、またはJSONのパースエラーや必須キーの欠落等による不正フォーマットを検知した場合、`DEFAULT_CONFIG` を用いて JSON ファイルを新規生成または**上書き**する。
  - 上書きを実施する際は「既存の設定ファイルが不正なため、初期設定で上書きしました」という旨のエラーログまたは警告ログを出力する。
  - `json.dump` の際に `ensure_ascii=False` および `encoding='utf-8'` を明示する。
  - `PermissionError` でファイルの生成・上書きに失敗した場合はエラーを捕捉し、オンメモリのデフォルト設定のまま例外を吐かず起動を継続させる。
- **`get_api_key(self)` 関数の修正**:
  - 環境変数の読み込み処理を削除し、`self.get("api.gemini_api_key")` からの取得に単一化する。
- **`get_model_options(self)` と `get_model_description(self, model_name)` の修正**:
  - 後方互換性処理は持たず、新構造（辞書のリスト）を前提として各データを抽出して返すように修正する。

### アプリケーションメイン

#### [MODIFY] `main.py`(file:///c:/auto_class/main.py)

- **環境変数読み込み処理の削除**:
  - `from dotenv import load_dotenv` のインポート文を削除する。
  - Gemini API 設定部分（`__init__` メソッド内）にある `load_dotenv()` の呼び出しを削除する。
- **UI連携**: `ConfigManager` 側で仕様変更に伴う戻り値の型（List[str] および str）が変わらないように設計するため、UIのプルダウン更新処理自体は既存のままで動作する。

### セキュリティおよびサンプルファイル設定

#### [MODIFY] `.gitignore`(file:///c:/auto_class/.gitignore)

- `config.json` を追記し、Gitの管理から除外する。

#### [NEW] `config.example.json`(file:///c:/auto_class/config.example.json)

- `DEFAULT_CONFIG` と同じ新しい設定内容を持つサンプルファイルを作成する（APIキーはダミー値）。

### ドキュメント変更

#### [MODIFY] `CONFIG.md`(file:///c:/auto_class/CONFIG.md)

- JSONの例を新しい `models` 構造に合わせるよう書き換える。
- APIキー設定に関して環境変数の記述を削除し、「`config.example.json` をコピーして設定する」旨に修正する。

## 検証計画

### 自動テスト

- **モックを用いた一時テスト**:
  1. `config.json` が存在しない環境で、正しく生成されることを確認する。
  2. 意図的に壊れた JSON パターンを用意し、実行時に**エラーログを出力しつつデフォルト値で上書きされる**ことを確認する。
  3. APIキー取得時に環境変数が無視されることを確認する。
- 確認後、一時スクリプトは消去する。
