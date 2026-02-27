# Configuration Management

この文書では、auto_class アプリケーションの設定管理システムについて説明します。

## 概要

アプリケーションでは、`config.json` ファイルで設定値を管理しています。

> [!TIP]
> 初回起動時など `config.json` が存在しない場合、デフォルト値を持つ設定ファイルが自動生成されます。

## 設定ファイルの準備

1. `config.example.json` をコピーして `config.json` を作成します。
2. `api.gemini_api_key` に自分のAPIキーを設定します。

```
cp config.example.json config.json
```

## 設定ファイル (`config.json`)

```json
{
  "api": {
    "gemini_api_key": "YOUR_GEMINI_API_KEY",
    "models": [
      {
        "model_name": "gemini-3-flash-preview",
        "description": "一般用途、無料枠"
      },
      {
        "model_name": "gemini-3.1-pro-preview",
        "description": "高性能モデル、有料枠のみ使用可"
      }
    ],
    "default_model_index": 0
  },
  "audio": {
    "no_sound_timeout_seconds": 180,
    "silence_threshold": 0.01
  },
  "screenshot": {
    "similarity_threshold": 0.83,
    "diff_pixel_threshold": 10
  },
  "ui": {
    "window_width": 600,
    "window_height": 550,
    "min_width": 600,
    "min_height": 550
  },
  "logging": {
    "level": "INFO",
    "filename": "slide_capture_app.log",
    "format": "%(asctime)s - %(levelname)s - %(threadName)s - %(message)s"
  }
}
```

## 設定項目の説明

### API設定 (`api`)

- `gemini_api_key`: Gemini APIキー（`config.json` から読み込む）
- `models`: 利用可能なGeminiモデルのリスト
  - `model_name`: モデル名
  - `description`: 用途・説明
- `default_model_index`: デフォルトで選択されるモデルのインデックス

### 音声設定 (`audio`)

- `no_sound_timeout_seconds`: 無音状態での録画終了タイムアウト（秒）
- `silence_threshold`: 無音と判定する振幅の閾値（0.0-1.0）

### スクリーンショット設定 (`screenshot`)

- `similarity_threshold`: 画像の類似度閾値
- `diff_pixel_threshold`: 差分ピクセルの閾値

### UI設定 (`ui`)

- `window_width`, `window_height`: アプリケーションウィンドウのサイズ
- `min_width`, `min_height`: ウィンドウの最小サイズ

### ログ設定 (`logging`)

- `level`: ログレベル（DEBUG, INFO, WARNING, ERROR, CRITICAL）
- `filename`: ログファイル名
- `format`: ログメッセージのフォーマット

## APIキーの設定

APIキーは `config.json` の `api.gemini_api_key` フィールドに設定してください。

> [!CAUTION]
> `config.json` には APIキーが含まれるため、`.gitignore` によってGit管理対象から除外されています。リポジトリにコミットしないよう注意してください。

## 設定の自動生成と上書き

アプリ起動時に以下の状況が検出された場合、`config.json` をデフォルト値で自動生成または上書きします。

- `config.json` が存在しない
- `config.json` が不正なJSON
- `config.json` が旧形式（モデルが文字列リスト形式）

上書き時はログに警告メッセージが出力されます。

## 設定の変更

1. `config.json` ファイルを編集
2. アプリケーションを再起動

モデル一覧を変更するには `models` 配列を更新してください（各要素に `model_name` と `description` が必要）。
