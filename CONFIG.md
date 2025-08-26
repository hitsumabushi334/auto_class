# Configuration Management

この文書では、auto_class アプリケーションの設定管理システムについて説明します。

## 概要

アプリケーションでは、以前はハードコードされていた設定値を `config.json` ファイルで管理するようになりました。これにより、モデルの更新、音声設定の調整、UI設定などが容易になります。

## 設定ファイル (`config.json`)

```json
{
    "api": {
        "gemini_api_key": "MOCK_API_KEY_FOR_DEVELOPMENT",
        "models": [
            "gemini-2.5-flash-preview-04-17",
            "gemini-2.0-flash"
        ],
        "default_model_index": 0,
        "model_descriptions": {
            "gemini-2.5-flash-preview-04-17": "用途: 短時間動画向け",
            "gemini-2.0-flash": "用途: 長時間動画向け"
        }
    },
    "audio": {
        "no_sound_timeout_seconds": 180,
        "silence_threshold": 0.01
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
- `gemini_api_key`: Gemini APIキー（デフォルトはモック値）
- `models`: 利用可能なGeminiモデルのリスト
- `default_model_index`: デフォルトで選択されるモデルのインデックス
- `model_descriptions`: 各モデルの説明文

### 音声設定 (`audio`)
- `no_sound_timeout_seconds`: 無音状態での録画終了タイムアウト（秒）
- `silence_threshold`: 無音と判定する振幅の閾値（0.0-1.0）

### UI設定 (`ui`)
- `window_width`, `window_height`: アプリケーションウィンドウのサイズ
- `min_width`, `min_height`: ウィンドウの最小サイズ

### ログ設定 (`logging`)
- `level`: ログレベル（DEBUG, INFO, WARNING, ERROR, CRITICAL）
- `filename`: ログファイル名
- `format`: ログメッセージのフォーマット

## APIキーの設定

APIキーは以下の優先順位で決定されます：

1. 環境変数 `GEMINI_API_KEY`
2. 設定ファイルの `api.gemini_api_key`

セキュリティのため、本番環境では環境変数を使用することを推奨します。

## 設定の変更

設定を変更するには：

1. `config.json` ファイルを編集
2. アプリケーションを再起動

モデル一覧を変更した場合は、`models` 配列と `model_descriptions` オブジェクトの両方を更新してください。

## モック値について

開発環境では `MOCK_API_KEY_FOR_DEVELOPMENT` がデフォルトのAPIキーとして設定されており、実際のAPI呼び出しは行われません。本番環境では有効なAPIキーを設定してください。