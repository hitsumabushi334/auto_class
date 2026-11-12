# auto-class

講義やプレゼンのウィンドウを録画し、変化のあった画面をスクリーンショットとして残し、録画終了後に Gemini API で構造化されたノートを作る Windows 向けデスクトップアプリです。ウィンドウタイトルは「スライドキャプチャ＆録画」です。

## できること

- 選択したウィンドウを約 1 FPS で録画し、既定のスピーカー出力をループバック録音して 1 本の MP4 にまとめます。
- 同じ領域を 1 秒間隔で見て、直前に保存した画像との差分が閾値を超えたときだけ PNG を保存します。ウィンドウ未選択時のスクリーンショットはプライマリモニターが対象です。
- 無音が `audio.no_sound_timeout_seconds`（既定 180 秒）を超えると録画を自動停止します。
- 録画の保存に成功すると、動画を Gemini に渡し、トピック・要点・専門用語・課題を含むノートを Markdown（既定）または Word で保存します。
- キャプチャ中は Windows のスリープとディスプレイオフを抑止します。Escape で停止できます。

画面録画とスクリーンショットはそれぞれオン/オフできます。両方オフのときは開始できません。録画には対象ウィンドウの選択が必要です。

## 動作環境

- Windows（`pywin32` によるウィンドウ列挙、DPI 設定、スリープ制御を使います）
- Python 3.13 以上
- [FFmpeg](https://ffmpeg.org/) が PATH 上にあること（MoviePy が MP4 書き出しに使います）
- ループバック録音ができる既定のスピーカー
- [Google AI Studio](https://aistudio.google.com/) の Gemini API キー（ノート生成に必要。録画とスクリーンショットだけなら空のままで起動できます）

## セットアップ

[uv](https://docs.astral.sh/uv/) を使う場合:

```bash
uv sync
```

pip を使う場合:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

初回起動時、`config.json` が無い・壊れている・モデル定義が旧形式（文字列の配列）のときは、`config_manager.py` の `DEFAULT_CONFIG` で作り直します。API キーは設定タブ、または生成された `config.json` の `api.gemini_api_key` に入れます。`config.example.json` はキーを入れるときのひな型です。

`config.json` は API キーを含むため `.gitignore` 対象です。リポジトリにコミットしません。

## 起動

```bash
uv run python main.py
```

## 使い方

1. Home タブで保存フォルダ名を入れます。空欄のときは作業ディレクトリ直下に `capture_YYYYMMDD_HHMMSS` を作ります。相対パスは作業ディレクトリ基準、絶対パスはそのまま使います。
2. 「更新」でウィンドウ一覧を読み、録画したいウィンドウを選びます。
3. 使う Gemini モデルを選びます。説明文は設定の `api.models` から表示されます。
4. 「画面録画」「スクリーンショット」の必要な方をオンにして「開始」を押します。実行中は設定タブ、フォルダ名、ウィンドウ選択、モデル選択をロックします。
5. 「停止」または Escape で終了します。録画していた場合は MP4 を書き出したあとノート生成が始まり、完了までウィンドウは閉じられません。スクリーンショットだけのときはノート生成は走りません。

設定タブでは API キー、モデル一覧、無音判定、スクリーンショットの差分閾値、ウィンドウサイズ、ログ、ノート形式、開発者モードを編集できます。「保存」で `config.json` と実行中の値へ反映し、「リセット」で既定値に戻します。処理の開始に失敗したときは設定タブのロックを戻します。

## 出力

保存フォルダには次のファイルが残ります。

| ファイル | 内容 |
| --- | --- |
| `recording_YYYYMMDD_HHMMSS.mp4` | ウィンドウ映像とループバック音声 |
| `note_recording_YYYYMMDD_HHMMSS.md` または `.docx` | Gemini の応答から作ったノート。形式は `app.note_output_mode` |
| `screenshot_YYYYMMDD_HHMMSS_mmm.png` | 画面変化を検出したときの静止画 |

フレームと音声の一時ファイル（`_temp_frames`、`_temp_audio.raw`）は動画の書き出し後に削除します。ログは既定で作業ディレクトリの `slide_capture_app.log` です。`app.developer_mode` をオンにすると同じ内容をコンソールにも出します。

ノートの構成は次のとおりです。

- タイトルと全体要約
- トピックごとの見出し（時間範囲）、キーワード、要約、ポイント
- 重要な知見（根拠の種別、重要度、タイムスタンプ）
- 専門用語（講義での定義、誤解しやすい点）
- 課題・提出物（期限とタイムスタンプ。配列が空なら節自体を出しません）

根拠の種別は `lecture_explicit`（講義で明示）、`lecture_inferred`（講義内容から推論）、`model_background`（一般知識による補完）です。

## 設定

項目の意味と読み込み規則の詳細は [CONFIG.md](CONFIG.md) にあります。現行の既定値は次の形です。`app` はノート形式と開発者モードです。

```json
{
  "api": {
    "gemini_api_key": "",
    "models": [
      { "model_name": "gemini-3-flash-preview", "description": "一般用途、無料枠" },
      { "model_name": "gemini-3.1-pro-preview", "description": "高性能モデル、有料枠のみ使用可" }
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
  },
  "app": {
    "note_output_mode": "MD",
    "developer_mode": false
  }
}
```

`screenshot.similarity_threshold` を上げるほど、わずかな差分でも別画像として保存します。`screenshot.diff_pixel_threshold` は差分と数える最小の画素差です。

## ディレクトリ

| パス | 役割 |
| --- | --- |
| `main.py` | Tkinter UI、録画、スクリーンショット、Gemini への動画送信、ノート保存 |
| `config_manager.py` | `config.json` の生成、検証、更新、リセット |
| `note_formatter.py` | 要約 JSON から Markdown と Word を組み立てる |
| `app_logging.py` | 起動時のログハンドラ設定 |
| `prompt.md` | ノート生成プロンプトの参照用。実行時は `main.py` 内の同内容を使います |
| `schema.json` | Gemini 応答スキーマの参照用。実行時は `main.py` 内の同内容を使います |
| `note_templete.md` | Markdown ノートのレイアウト仕様。実装は `note_formatter.py` |
| `gemini_video_summary_test.py` | 既存の動画から Word ノートを試す単体スクリプト |
| `config.example.json` | API キー記入用のひな型 |
| `tests/` | `config_manager`、`note_formatter`、ログ初期化のテスト |
| `CONFIG.md` | 設定項目の説明 |
| `CHANGELOG.md` | 変更履歴 |

`gemini_video_summary_test.py` は GUI とは別の短いプロンプトで Word を出します。

```bash
uv run python gemini_video_summary_test.py path\to\video.mp4 --output path\to\note.docx
```

## テスト

```bash
uv run pytest
```

GUI の録画・キャプチャ・API 呼び出しは自動テストの対象外です。設定の読み書き、ノートの整形、起動時ログを検証します。

## 補足

- ノート生成は選択中のモデルで動画をアップロードし、JSON スキーマ付きで応答を受け取ります。失敗時は間隔を空けて再試行し、完了後にアップロードしたファイルを削除します。
- モデル名や無料枠の条件は Gemini 側の変更に従います。使える名前は設定タブのモデルリストで差し替えます。
- 変更履歴は [CHANGELOG.md](CHANGELOG.md) を見てください。
