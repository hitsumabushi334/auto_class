# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
from datetime import datetime
import os
import logging  # logging モジュールをインポート
import traceback  # スタックトレース取得のため
from ctypes import windll


import subprocess
import win32gui  # ウィンドウ操作のため追加
from google import genai
from docx import Document  # mdファイル生成のため追加
from docx.shared import Inches  # mdファイル生成のため追加 (必要に応じて)
import json
from markd import Markdown

# import sounddevice as sd # sounddevice を削除
import queue  # 音声データキューのため追加

# import pyaudio  # PyAudio を削除
import soundcard as sc  # SoundCard を追加

# import ffmpeg  # ffmpeg-python は削除
import moviepy.editor as mpe  # MoviePy を追加
from moviepy.audio.AudioClip import AudioArrayClip  # AudioArrayClip を直接インポート

import mss  # 画面キャプチャのため追加
import cv2
from PIL import Image, UnidentifiedImageError
import numpy as np

from config_manager import get_config_manager
from note_formatter import populate_word_document, render_markdown_note
from pathlib import Path
from app_logging import configure_initial_logging

# --- ロギング設定 ---
config_manager = get_config_manager()
logging_settings = config_manager.get_logging_settings()
app_settings = config_manager.get_app_settings()
configure_initial_logging(
    logging_settings=logging_settings,
    developer_mode=app_settings["developer_mode"],
)
logger = logging.getLogger(__name__)


def resource_path(*parts: str) -> Path:
    base_dir = Path(__file__).resolve().parent
    return base_dir.joinpath(*parts)


class SlideCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("スライドキャプチャ＆録画")

        # Load configuration
        self.config = get_config_manager()
        ui_settings = self.config.get_ui_settings()

        # UIの高さを増やして新しい要素を配置
        self.root.geometry(
            f"{ui_settings['window_width']}x{ui_settings['window_height']}"
        )  # サイズ調整
        self.root.update_idletasks()
        self.root.minsize(ui_settings["min_width"], ui_settings["min_height"])

        # --- 状態変数 ---
        self.recording_enabled = True  # 録画アクティブフラグ
        self.capturing_enabled = True  # キャプチャアクティブフラグ
        self.is_capturing_screenshot = False  # スクリーンショット中フラグ
        self.is_recording = False  # 録画中フラグ
        self.screenshot_thread = None
        self.recording_thread = None
        self.audio_recording_thread = None  # 音声録音スレッド
        self.recorded_frames = []  # 録画フレームを一時保存するリスト
        self.audio_queue = queue.Queue()  # 音声データを一時保存するキュー
        self.last_screenshot_image = (
            None  # キャプチャした最新の画像（保存されたかどうかに関わらず）
        )
        self.last_saved_screenshot_image = None  # 最後に保存した画像（比較用）
        self.screenshot_saved_count = 0
        self.last_saved_screenshot_filename = ""
        self.recording_start_time = None
        self.last_saved_video_filename = ""
        self.save_folder_name = tk.StringVar()
        self.selected_window_handle = tk.IntVar(value=0)  # 選択されたウィンドウハンドル
        self.error_occurred_in_screenshot_thread = False
        self.error_occurred_in_recording_thread = False
        self.note_creation_status = tk.StringVar(value="")  # ノート作成ステータス
        self.note_result_message = tk.StringVar(value="")  # ノート作成結果
        self.gemini_client = None  # Gemini API クライアント
        self.audio_sample_rate = None  # SoundCard で取得したサンプルレートを保存
        self.audio_channels = None  # SoundCard で取得したチャンネル数を保存
        self.last_sound_time = None  # 最後に音声を検知した時刻

        # Load app settings (note output mode / developer mode)
        app_settings = self.config.get_app_settings()
        self.note_output_mode = app_settings["note_output_mode"]  # "MD" or "Word"
        self.developer_mode = app_settings["developer_mode"]

        # Load audio settings from configuration
        audio_settings = self.config.get_audio_settings()
        self.no_sound_timeout_seconds = audio_settings[
            "no_sound_timeout_seconds"
        ]  # 無音状態のタイムアウト秒数
        self.silence_threshold = audio_settings[
            "silence_threshold"
        ]  # 無音と判定する振幅の閾値

        # Load screenshot settings from configuration
        screenshot_settings = self.config.get_screenshot_settings()
        self.similarity_threshold = screenshot_settings[
            "similarity_threshold"
        ]  # 画像の類似度判定の閾値
        self.diff_pixel_threshold = screenshot_settings[
            "diff_pixel_threshold"
        ]  # 差分ピクセルの閾値

        # Load model options from configuration
        self.gemini_model_options = self.config.get_model_options()

        self.selected_gemini_model = tk.StringVar(
            value=self.config.get_default_model()  # デフォルトモデルを設定から取得
        )

        # --- Gemini API 設定 ---
        try:
            api_key = self.config.get_api_key()  # 設定から API キーを取得
            if not api_key or api_key == "MOCK_API_KEY_FOR_DEVELOPMENT":
                logger.warning(
                    "APIキーが設定されていないか、モックキーが使用されています。ノート作成機能は利用できません。"
                )
                messagebox.showwarning(
                    "APIキー未設定",
                    "APIキーが設定されていないか、モックキーが使用されています。\nノート作成機能は利用できません。",
                )
            else:
                # 動画を扱えるモデルを指定 (例: gemini-1.5-pro-latest)
                # 注意: モデル名や利用可否は変更される可能性があるため、ドキュメントを確認すること
                # self.gemini_model = genai.GenerativeModel('gemini-1.5-pro-latest') # 後で使う
                logger.info("Gemini API クライアントを設定しました。")
                self.gemini_client = genai.Client(api_key=api_key)
        except Exception as e:
            logger.exception("Gemini API の設定中にエラーが発生しました。")
            messagebox.showerror(
                "API設定エラー", f"Gemini API の設定に失敗しました:\n{e}"
            )

        # --- UI要素の作成 ---

        # --- タブ構成: ttk.Notebook ---
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        home_tab = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(home_tab, text="Home")

        # Settings タブは _build_settings_tab() で別途追加
        self._build_settings_tab()

        # --- 保存フォルダ設定 ---
        folder_frame = ttk.Frame(home_tab, padding="10")
        folder_frame.pack(fill=tk.X)
        folder_label = ttk.Label(folder_frame, text="保存フォルダ名:")
        folder_label.pack(side=tk.LEFT, padx=(0, 5))
        self.folder_entry = ttk.Entry(
            folder_frame, textvariable=self.save_folder_name, width=45  # 幅調整
        )
        self.folder_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # --- ウィンドウ選択 ---
        window_selection_frame = ttk.LabelFrame(
            home_tab, text="録画対象ウィンドウ選択", padding="10"
        )
        window_selection_frame.pack(fill=tk.X, padx=10, pady=5)

        self.window_listbox = tk.Listbox(
            window_selection_frame, height=5, exportselection=False
        )
        self.window_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        self.window_listbox.bind("<<ListboxSelect>>", self.on_window_select)  # 実装

        window_scrollbar = ttk.Scrollbar(
            window_selection_frame,
            orient=tk.VERTICAL,
            command=self.window_listbox.yview,
        )
        window_scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.window_listbox.config(yscrollcommand=window_scrollbar.set)

        self.refresh_window_list_button = ttk.Button(
            window_selection_frame,
            text="更新",
            command=self.refresh_window_list,  # あとで実装
        )
        self.refresh_window_list_button.pack(side=tk.LEFT, padx=5)

        # --- Gemini モデル選択 ---
        model_selection_frame = ttk.LabelFrame(
            home_tab, text="Gemini モデル選択", padding="10"
        )
        model_selection_frame.pack(fill=tk.X, padx=10, pady=5)

        model_label = ttk.Label(model_selection_frame, text="モデル:")
        model_label.pack(side=tk.LEFT, padx=(0, 5))

        self.model_combobox = ttk.Combobox(
            model_selection_frame,
            textvariable=self.selected_gemini_model,
            values=self.gemini_model_options,
            state="readonly",  # ユーザー入力不可
            width=40,  # 幅調整
        )
        self.model_combobox.pack(side=tk.LEFT, padx=(0, 5))  # 右に少し余白
        self.model_combobox.bind(
            "<<ComboboxSelected>>", self._on_gemini_model_selected
        )  # イベントハンドラをバインド

        # 用途表示ラベルを追加（モデル選択フレームの下に配置）
        model_description_frame = ttk.Frame(home_tab, padding=(10, 0, 10, 5))
        model_description_frame.pack(fill=tk.X)
        self.gemini_model_description_label = ttk.Label(
            model_description_frame,
            text="",
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=500,
        )
        self.gemini_model_description_label.pack(fill=tk.X)

        # --- 操作ボタン (統合) ---
        button_frame = ttk.Frame(home_tab, padding="10")
        button_frame.pack(fill=tk.X)

        self.start_button = ttk.Button(
            button_frame,
            text="開始",
            command=self.start_tasks,  # 新しい統合開始メソッド
        )
        self.start_button.pack(side=tk.LEFT, padx=5)
        self.stop_button = ttk.Button(
            button_frame,
            text="停止",
            command=self.stop_all_tasks,
            state=tk.DISABLED,  # 統合停止メソッド
        )
        self.stop_button.pack(side=tk.LEFT, padx=5)

        # 録画有効/無効チェックボックス
        self.recording_enabled_var = tk.BooleanVar(value=self.recording_enabled)
        self.recording_checkbox = ttk.Checkbutton(
            button_frame,
            text="画面録画",
            variable=self.recording_enabled_var,
            command=self.on_toggle_recording,
        )
        self.recording_checkbox.pack(side=tk.LEFT, padx=5)

        # スクショット有効/無効チェックボックス (将来の拡張用)
        self.capturing_enabled_var = tk.BooleanVar(value=self.capturing_enabled)
        self.capturing_checkbox = ttk.Checkbutton(
            button_frame,
            text="スクリーンショット",
            variable=self.capturing_enabled_var,
            command=self.on_toggle_capturing,
        )
        self.capturing_checkbox.pack(side=tk.LEFT, padx=5)

        # --- ステータス表示 ---
        status_frame = ttk.Frame(home_tab, padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True)

        # スクリーンショットステータス
        screenshot_status_frame = ttk.LabelFrame(
            status_frame, text="スクリーンショット", padding="5"
        )
        screenshot_status_frame.pack(fill=tk.X, pady=(0, 5))
        self.screenshot_status_label = ttk.Label(
            screenshot_status_frame,
            text="待機中...",
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=400,
        )
        self.screenshot_status_label.pack(fill=tk.X)

        # 録画ステータス
        recording_status_frame = ttk.LabelFrame(status_frame, text="録画", padding="5")
        recording_status_frame.pack(fill=tk.X, pady=(0, 5))
        self.recording_status_label = ttk.Label(
            recording_status_frame,
            text="待機中...",
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=400,
        )
        self.recording_status_label.pack(fill=tk.X)

        # ノート作成ステータス
        note_status_frame = ttk.LabelFrame(status_frame, text="ノート作成", padding="5")
        note_status_frame.pack(fill=tk.X, pady=(0, 5))
        self.note_status_label = ttk.Label(
            note_status_frame,
            textvariable=self.note_creation_status,
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=400,
        )
        self.note_status_label.pack(fill=tk.X)
        self.note_result_label = ttk.Label(
            note_status_frame,
            textvariable=self.note_result_message,
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=400,
            foreground="blue",
        )
        self.note_result_label.pack(fill=tk.X)

        # --- イベントバインド ---
        self.root.bind("<Escape>", lambda e: self.stop_all_tasks())  # Escで全停止
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.original_on_closing = self.on_closing  # 元の閉じる処理を保持

        logger.info("アプリケーションを初期化しました。")
        self.refresh_window_list()  # 初期ウィンドウリスト表示
        self._on_gemini_model_selected()  # 初期モデルの用途を表示

        # 起動時に developer_mode を適用 (StreamHandler を除去 or 維持)
        self._apply_settings_to_app(self.config.get_current_config())

        # 設定ファイルの自動生成・上書きが発生した場合、UI で通知する
        if self.config.config_was_auto_generated:
            messagebox.showwarning(
                "設定ファイルの初期化",
                f"設定ファイルを初期化しました。\n\n"
                f"理由: {self.config.config_auto_generated_reason}\n\n"
                f"デフォルト設定で動作します。\n"
                f"APIキーを設定するには {self.config.config_path} を開いて\n"
                f"api.gemini_api_key に APIキーを入力してください。",
            )

    # --- ウィンドウ選択関連メソッド ---
    # ----------------------------------------------------------------------- #
    # 設定タブ構築
    # ----------------------------------------------------------------------- #

    def _build_settings_tab(self):
        """設定タブを構築し、self.notebook に追加する。"""
        settings_tab = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(settings_tab, text="設定")
        self._settings_tab_frame = (
            settings_tab  # タブ自体は Frame として参照不要、indexで操作
        )

        # スクロール可能エリア
        canvas_frame = ttk.Frame(settings_tab)
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(canvas_frame)
        scrollbar = ttk.Scrollbar(
            canvas_frame, orient=tk.VERTICAL, command=canvas.yview
        )
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        inner_frame = ttk.Frame(canvas, padding=10)
        inner_window = canvas.create_window((0, 0), window=inner_frame, anchor=tk.NW)

        def _on_inner_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            # inner_frameの幅をCanvasの幅に合わせる（スクロールバー分を引く）
            canvas.itemconfig(inner_window, width=event.width)

        inner_frame.bind("<Configure>", _on_inner_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # マウスホイールでスクロール
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self._build_settings_widgets(inner_frame)

        # 下部: 保存・リセットボタン (中央配置)
        btn_frame = ttk.Frame(settings_tab, padding=(0, 10))
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X)
        btn_inner = ttk.Frame(btn_frame)
        btn_inner.pack(anchor=tk.CENTER)
        ttk.Button(
            btn_inner, text="保存", command=self._on_settings_save, width=15
        ).pack(side=tk.LEFT, padx=10)
        ttk.Button(
            btn_inner, text="リセット", command=self._on_settings_reset, width=15
        ).pack(side=tk.LEFT, padx=10)

    def _build_settings_widgets(self, parent):
        """設定フォームのウィジェット群を構築する（Gridレイアウト）。"""
        self._settings_vars = {}  # key: "section.key" -> tk.Variable

        # column 0: Label, 1: Widget, 2: Description
        parent.columnconfigure(0, weight=0, minsize=160)
        parent.columnconfigure(1, weight=0, minsize=80)
        parent.columnconfigure(2, weight=1)  # 説明文が伸びるように

        self._current_row = 0

        def section(text):
            ttk.Separator(parent, orient=tk.HORIZONTAL).grid(
                row=self._current_row, column=0, columnspan=3, sticky="we", pady=(15, 5)
            )
            self._current_row += 1
            ttk.Label(parent, text=text, font=("", 10, "bold")).grid(
                row=self._current_row,
                column=0,
                columnspan=3,
                sticky="w",
                padx=5,
                pady=(0, 5),
            )
            self._current_row += 1

        def row_field(
            key,
            label_text,
            description,
            widget_type="entry",
            options: list[str] | tuple[str, ...] = (),
        ):
            ttk.Label(parent, text=label_text).grid(
                row=self._current_row, column=0, sticky="w", padx=5, pady=5
            )

            widget_frame = ttk.Frame(parent)
            widget_frame.grid(
                row=self._current_row, column=1, sticky="w", padx=5, pady=5
            )

            var = None
            if widget_type == "entry":
                var = tk.StringVar(value=str(self.config.get(key, "")))
                ttk.Entry(widget_frame, textvariable=var, width=15).pack(fill=tk.X)
            elif widget_type == "combobox":
                var = tk.StringVar(value=str(self.config.get(key, "")))
                ttk.Combobox(
                    widget_frame,
                    textvariable=var,
                    values=options,
                    state="readonly",
                    width=13,
                ).pack(fill=tk.X)
            elif widget_type == "bool":
                var = tk.BooleanVar(value=bool(self.config.get(key, False)))
                ttk.Checkbutton(widget_frame, variable=var).pack(side=tk.LEFT)

            if var is not None:
                self._settings_vars[key] = var

            desc_label = ttk.Label(parent, text=description, foreground="gray")
            desc_label.grid(
                row=self._current_row, column=2, sticky="we", padx=5, pady=5
            )
            # 見切れを防ぐため、実際の幅から少しマージン（-10px）を引いた値をwraplengthに設定する
            desc_label.bind(
                "<Configure>",
                lambda e, l=desc_label: l.config(wraplength=max(50, e.width - 10)),
            )

            self._current_row += 1

        # --- API ---
        section("API設定")
        row_field(
            "api.gemini_api_key",
            "Gemini APIキー",
            "Google AI Studio から取得したAPIキー。",
        )
        row_field(
            "api.default_model_index",
            "デフォルトモデル (ｲﾝﾃﾞｯｸｽ)",
            "使用するモデルリストのインデックス (0始まり)。",
        )

        # models の動的エディタ
        self._build_models_editor(parent)

        # --- 音声 ---
        section("音声設定")
        row_field(
            "audio.no_sound_timeout_seconds",
            "無音タイムアウト (秒)",
            "この秒数を超えて無音が続くと録画を停止します。大きくすると長い無音を許容。",
        )
        row_field(
            "audio.silence_threshold",
            "無音判定閾値",
            "この振幅以下の音声を「無音」と判定します。小さくするとより小さな音を拾います。",
        )

        # --- スクリーンショット ---
        section("スクリーンショット設定")
        row_field(
            "screenshot.similarity_threshold",
            "類似度閾値",
            "0〜1の値。大きくするほど判定が厳しくなり、差分が小さい変化を保存しやすくなります。",
        )
        row_field(
            "screenshot.diff_pixel_threshold",
            "差分ピクセル閾値",
            "差分として扱う最小ピクセル数。小さくすると微細な変化も保存します。",
        )

        # --- UI ---
        section("ウィンドウ設定")
        row_field("ui.window_width", "ウィンドウ幅", "アプリウィンドウの幅 (px)。")
        row_field("ui.window_height", "ウィンドウ高さ", "アプリウィンドウの高さ (px)。")
        row_field("ui.min_width", "最小幅", "ウィンドウの最小幅 (px)。")
        row_field("ui.min_height", "最小高さ", "ウィンドウの最小高さ (px)。")

        # --- ログ ---
        section("ログ設定")
        row_field(
            "logging.level",
            "ログレベル",
            "出力するログの深刻度を指定します。DEBUGは全詳細、INFOは通常動作、WARNINGは警告のみを出力します。",
            widget_type="combobox",
            options=["DEBUG", "INFO", "WARNING", "ERROR"],
        )
        row_field("logging.filename", "ログファイル名", "ログを保存するファイル名。")
        row_field("logging.format", "ログフォーマット", "ログ行のフォーマット文字列。")

        # --- アプリ ---
        section("アプリ設定")
        row_field(
            "app.note_output_mode",
            "ノート出力形式",
            "MD: Markdownファイル (.md) / Word: Word文書 (.docx) を選択。",
            widget_type="combobox",
            options=["MD", "Word"],
        )
        row_field(
            "app.developer_mode",
            "開発者モード",
            "オンにするとコンソール (ターミナル) にもログが出力されます。",
            widget_type="bool",
        )

    def _build_models_editor(self, parent):
        """api.models を動的に編集するリストUIを構築する。"""
        self.models_container = ttk.Frame(parent)
        self.models_container.grid(
            row=self._current_row, column=0, columnspan=3, sticky="we", padx=5, pady=5
        )
        self._current_row += 1

        header_frame = ttk.Frame(self.models_container)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header_frame, text="モデルリスト:").pack(side=tk.LEFT)
        ttk.Label(
            header_frame, text="モデル名と説明のペアを定義します。", foreground="gray"
        ).pack(side=tk.LEFT, padx=5)

        self.models_list_frame = ttk.Frame(self.models_container)
        self.models_list_frame.pack(fill=tk.BOTH, expand=True)

        self._model_entries = []  # type: list[dict]

        self._refresh_models_list_ui()

        ttk.Button(
            self.models_container,
            text="＋ モデルを追加",
            width=15,
            command=lambda: self._add_model_row(),
        ).pack(anchor="w", pady=5)

    def _refresh_models_list_ui(self):
        """models_list_frame の中身を現在の設定から再構築する。"""
        # 既存のエントリをクリア
        for widget in self.models_list_frame.winfo_children():
            widget.destroy()
        self._model_entries.clear()

        current_models = self.config.get("api.models", [])
        for m in current_models:
            m_name = m.get("model_name", "")
            m_desc = m.get("description", "")
            self._add_model_row(m_name, m_desc)

    def _add_model_row(self, model_name="", description=""):
        """モデル追加UIの一行を追加する。"""
        row_frame = ttk.Frame(self.models_list_frame)
        row_frame.pack(fill=tk.X, pady=2)

        ttk.Label(row_frame, text="名前:").pack(side=tk.LEFT)
        name_var = tk.StringVar(value=model_name)
        ttk.Entry(row_frame, textvariable=name_var, width=15).pack(
            side=tk.LEFT, padx=(2, 10)
        )

        ttk.Label(row_frame, text="説明:").pack(side=tk.LEFT)
        desc_var = tk.StringVar(value=description)
        ttk.Entry(row_frame, textvariable=desc_var, width=30).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 5)
        )

        entry_dict = {"name": name_var, "desc": desc_var, "frame": row_frame}
        self._model_entries.append(entry_dict)

        def remove_self():
            row_frame.destroy()
            if entry_dict in self._model_entries:
                self._model_entries.remove(entry_dict)

        ttk.Button(row_frame, text="－", width=3, command=remove_self).pack(
            side=tk.LEFT
        )

    def _apply_settings_to_app(self, new_config: dict):
        """保存した設定を self に即時反映する。"""
        import logging as _logging

        # audio
        audio = new_config.get("audio", {})
        self.no_sound_timeout_seconds = audio.get(
            "no_sound_timeout_seconds", self.no_sound_timeout_seconds
        )
        self.silence_threshold = audio.get("silence_threshold", self.silence_threshold)

        # screenshot
        screenshot = new_config.get("screenshot", {})
        self.similarity_threshold = screenshot.get(
            "similarity_threshold", self.similarity_threshold
        )
        self.diff_pixel_threshold = screenshot.get(
            "diff_pixel_threshold", self.diff_pixel_threshold
        )

        # ui (window geometry)
        ui = new_config.get("ui", {})
        w = ui.get("window_width", self.config.get("ui.window_width", 600))
        h = ui.get("window_height", self.config.get("ui.window_height", 550))
        mw = ui.get("min_width", self.config.get("ui.min_width", 600))
        mh = ui.get("min_height", self.config.get("ui.min_height", 550))
        self.root.geometry(f"{w}x{h}")
        self.root.minsize(mw, mh)

        # app settings
        app = new_config.get("app", {})
        self.note_output_mode = app.get("note_output_mode", "MD")
        self.developer_mode = app.get("developer_mode", False)

        # logging: レベル・フォーマット・ハンドラの動的切り替え (P3 修正)
        logging_settings = new_config.get("logging", {})
        root_logger = _logging.getLogger()

        # 1. ログレベルの更新
        level_str = logging_settings.get("level", "INFO").upper()
        root_logger.setLevel(getattr(_logging, level_str, _logging.INFO))

        # 2. フォーマットの更新
        format_str = logging_settings.get(
            "format", "%(asctime)s - %(levelname)s - %(threadName)s - %(message)s"
        )
        formatter = _logging.Formatter(format_str)

        # 3. ハンドラの更新 (FileHandler と StreamHandler)
        file_filename = logging_settings.get("filename", "slide_capture_app.log")
        has_stream = False
        has_file = False

        # 既存ハンドラの更新・削除
        handlers_to_remove = []
        for h in root_logger.handlers:
            if isinstance(h, _logging.FileHandler):
                # ファイル名が変わった場合は古いものを削除
                if h.baseFilename != os.path.abspath(file_filename):
                    handlers_to_remove.append(h)
                else:
                    h.setFormatter(formatter)
                    has_file = True
            elif isinstance(h, _logging.StreamHandler):
                if self.developer_mode:
                    h.setFormatter(formatter)
                    has_stream = True
                else:
                    handlers_to_remove.append(h)

        for h in handlers_to_remove:
            root_logger.removeHandler(h)
            try:
                # uvコマンド等のラッパーによるsys.stderrの永久クローズを防ぐため、
                # FileHandler 以外は close() を呼ばないようにする
                if isinstance(h, _logging.FileHandler):
                    h.close()
            except Exception:
                pass

        # 不足ハンドラの追加
        if not has_file:
            fh = _logging.FileHandler(file_filename, encoding="utf-8")
            fh.setFormatter(formatter)
            root_logger.addHandler(fh)
        if self.developer_mode and not has_stream:
            import sys

            sh = _logging.StreamHandler(sys.stdout)
            sh.setFormatter(formatter)
            root_logger.addHandler(sh)

        # gemini model (Comboboxも更新)
        api = new_config.get("api", {})
        self.gemini_model_options = [
            m["model_name"]
            for m in api.get("models", [])
            if isinstance(m, dict) and "model_name" in m
        ]
        self.model_combobox["values"] = self.gemini_model_options

        # P2: アクティブなモデル選択を更新する
        current_selection = self.selected_gemini_model.get()
        if self.gemini_model_options:
            if current_selection not in self.gemini_model_options:
                # 現在の選択がリストから消えた場合はデフォルト(または先頭)に戻す
                default_idx = api.get("default_model_index", 0)
                if default_idx >= len(self.gemini_model_options):
                    default_idx = 0
                self.selected_gemini_model.set(self.gemini_model_options[default_idx])
                self._on_gemini_model_selected()  # UI説明文を更新
        else:
            self.selected_gemini_model.set("")
            self.gemini_model_description_label.config(text="用途: -")

        # P2: APIキーが変わった（またはクリアされた）場合、クライアントを再初期化・クリア
        api_key = api.get("gemini_api_key", "").strip()
        if (
            not api_key
            or api_key == "YOUR_GEMINI_API_KEY_HERE"
            or api_key == "MOCK_API_KEY_FOR_DEVELOPMENT"
        ):
            self.gemini_client = None
            logger.info(
                "APIキーが無効または未設定のため、Gemini APIクライアントをクリアしました。"
            )
        else:
            try:
                self.gemini_client = genai.Client(api_key=api_key)
                logger.info("Gemini APIクライアントを再初期化しました。")
            except Exception as e:
                logger.error(f"Gemini APIクライアントの再初期化に失敗しました: {e}")
                self.gemini_client = None

    def _on_settings_save(self):
        """設定保存ボタンの処理。バリデーション -> ConfigManager 更新 -> 即時反映。"""
        import json as _json

        new_config = {}
        errors = []
        for key, var in self._settings_vars.items():
            parts = key.split(".")
            section_key, field_key = parts[0], parts[1]

            raw = var.get()
            # 型変換
            default = self.config.get(key, None)
            if isinstance(default, bool):
                value = (
                    bool(var.get())
                    if isinstance(var, tk.BooleanVar)
                    else (raw.lower() == "true")
                )
            elif isinstance(default, int):
                try:
                    value = int(raw)
                except ValueError:
                    errors.append(f"{key}: 整数値を入力してください (入力値: {raw})")
                    continue
            elif isinstance(default, float):
                try:
                    value = float(raw)
                except ValueError:
                    errors.append(f"{key}: 数値を入力してください (入力値: {raw})")
                    continue
            else:
                value = raw

            if section_key not in new_config:
                new_config[section_key] = {}
            new_config[section_key][field_key] = value

        # --- api.models の読み取りとバリデーション ---
        models_data = []
        for entry in self._model_entries:
            m_name = entry["name"].get().strip()
            m_desc = entry["desc"].get().strip()
            if m_name:
                models_data.append({"model_name": m_name, "description": m_desc})

        if not models_data:
            errors.append("api.models: 1件以上のモデル名を設定してください。")
        else:
            if "api" not in new_config:
                new_config["api"] = {}
            new_config["api"]["models"] = models_data

            default_idx = new_config["api"].get("default_model_index", 0)
            if (
                not isinstance(default_idx, int)
                or default_idx < 0
                or default_idx >= len(models_data)
            ):
                errors.append(
                    f"api.default_model_index: 0 から {len(models_data)-1} の間の整数値を入力してください。"
                )

        if errors:
            from tkinter import messagebox as _mb

            _mb.showerror("入力エラー", "\n".join(errors))
            return

        # 変更前設定をマージ（設定外のキーを保持）
        import copy as _copy

        merged = _copy.deepcopy(self.config.get_current_config())
        for section_key, fields in new_config.items():
            if section_key not in merged:
                merged[section_key] = {}
            merged[section_key].update(fields)

        # ConfigManager 経由で保存（インメモリ + ファイル同時更新）
        self.config.update_config(merged)

        # main.py の各状態変数に即時反映
        self._apply_settings_to_app(merged)

        from tkinter import messagebox as _mb

        _mb.showinfo("保存完了", "設定を保存し、即時反映しました。")

    def _on_settings_reset(self):
        """リセットボタンの処理。ConfigManager をデフォルトに戻し、UIと状態変数を更新。"""
        from tkinter import messagebox as _mb

        if not _mb.askyesno("リセット確認", "すべての設定をデフォルト値に戻しますか？"):
            return

        self.config.reset_to_default()
        self._apply_settings_to_app(self.config.get_current_config())

        # ウィジェットの表示値を更新
        for key, var in self._settings_vars.items():
            current_val = self.config.get(key, "")
            if isinstance(var, tk.BooleanVar):
                var.set(bool(current_val))
            else:
                var.set(str(current_val))

        # モデルリストUIを再構築
        self._refresh_models_list_ui()

        _mb.showinfo("リセット完了", "設定をデフォルト値に戻しました。")

    # 設定タブのロック/アンロック
    def _lock_settings_tab(self):
        """録画/処理中に設定タブを無効化する。"""
        try:
            self.notebook.tab(1, state="disabled")
        except tk.TclError:
            pass

    def _unlock_settings_tab(self):
        """録画/処理完了後に設定タブを有効化する。"""
        try:
            self.notebook.tab(1, state="normal")
        except tk.TclError:
            pass

    def refresh_window_list(self):
        """実行中のウィンドウリストを取得し、リストボックスを更新する"""
        logger.info("ウィンドウリストを更新します。")
        self.window_listbox.delete(0, tk.END)
        self.window_handles = {}  # タイトルとハンドルのマッピングを保持

        try:

            def enum_windows_proc(hwnd, lParam):
                if win32gui.IsWindowVisible(hwnd):
                    text = win32gui.GetWindowText(hwnd)
                    if text:  # タイトルがあるウィンドウのみ
                        # 同じタイトルのウィンドウがある場合、ハンドルを追記して区別
                        display_text = f"{text} (HWND: {hwnd})"
                        self.window_listbox.insert(tk.END, display_text)
                        self.window_handles[display_text] = hwnd
                return True  # 列挙を続ける

            win32gui.EnumWindows(enum_windows_proc, None)
            logger.info(
                f"{self.window_listbox.size()} 個のウィンドウが見つかりました。"
            )
        except Exception as e:
            logger.exception("ウィンドウリストの取得中にエラーが発生しました。")
            messagebox.showerror(
                "エラー", f"ウィンドウリストの取得に失敗しました:\n{e}"
            )

        # self.start_recording_button.config(state=tk.DISABLED) # 個別ボタン削除

    def on_window_select(self, event):
        """リストボックスでウィンドウが選択されたときの処理"""
        selected_indices = self.window_listbox.curselection()
        if not selected_indices:
            self.selected_window_handle.set(0)
            # self.start_recording_button.config(state=tk.DISABLED) # 個別ボタン削除
            logger.info("ウィンドウ選択が解除されました。")
            return

        selected_index = selected_indices[0]
        selected_text = self.window_listbox.get(selected_index)

        # ハンドルを取得
        hwnd = self.window_handles.get(selected_text)
        if hwnd:
            self.selected_window_handle.set(hwnd)
            # self.start_recording_button.config(state=tk.NORMAL) # 個別ボタン削除
            logger.info(f"ウィンドウが選択されました: {selected_text} (HWND: {hwnd})")
            # 開始ボタンの状態は start_tasks で制御
        else:
            logger.error(
                f"選択されたテキストに対応するハンドルが見つかりません: {selected_text}"
            )
            self.selected_window_handle.set(0)
            # self.start_recording_button.config(state=tk.DISABLED) # 個別ボタン削除

    def _on_gemini_model_selected(self, event=None):
        """Geminiモデル選択コンボボックスの値が変更されたときに呼び出される"""
        selected_model = self.selected_gemini_model.get()
        description = f"用途: {self.config.get_model_description(selected_model)}"
        if description == "用途: 不明":
            logger.warning(f"不明なGeminiモデルが選択されました: {selected_model}")

        self.gemini_model_description_label.config(text=description)
        logger.info(
            f"Geminiモデル '{selected_model}' が選択されました。用途: {description}"
        )

    def on_toggle_recording(self):
        """録画有効/無効チェックボックスの状態が変更されたときに呼び出される"""
        self.recording_enabled = self.recording_enabled_var.get()
        logger.info(f"録画有効状態が変更されました: {self.recording_enabled}")
        if not self.recording_enabled and not self.is_recording:
            self.recording_status_label.config(
                text="待機中... (録画は無効化されています)"
            )
        elif not self.is_recording:
            self.recording_status_label.config(text="待機中...")

        self._update_start_button_state()

    def on_toggle_capturing(self):
        """スクリーンショット有効/無効チェックボックスの状態が変更されたときに呼び出される"""
        self.capturing_enabled = self.capturing_enabled_var.get()
        logger.info(
            f"スクリーンショット有効状態が変更されました: {self.capturing_enabled}"
        )
        if not self.capturing_enabled and not self.is_capturing_screenshot:
            self.screenshot_status_label.config(
                text="待機中... (スクリーンショットは無効化されています)"
            )
        elif not self.is_capturing_screenshot:
            self.screenshot_status_label.config(text="待機中...")

        self._update_start_button_state()

    def _update_start_button_state(self):
        """チェックボックスの状態に応じて開始ボタンの有効・無効を切り替える"""
        if self.is_capturing_screenshot or self.is_recording:
            return  # 実行中は start_tasks / stop_all_tasks に任せる

        if not self.recording_enabled and not self.capturing_enabled:
            self.start_button.config(state=tk.DISABLED)
        else:
            self.start_button.config(state=tk.NORMAL)

    def _set_checkbox_state(self, state):
        """録画・スクショのチェックボックスの状態をまとめて変更する"""
        self.recording_checkbox.config(state=state)
        self.capturing_checkbox.config(state=state)

    # --- 統合開始・停止メソッド ---
    def start_tasks(self):
        """スクリーンショットと録画を開始する（ウィンドウ選択状態による）"""
        if self.is_capturing_screenshot or self.is_recording:
            logger.warning("既にタスクが実行中です。")
            return

        # 録画/処理中は設定タブを無効化
        self._lock_settings_tab()

        hwnd = self.selected_window_handle.get()

        if not self.recording_enabled and not self.capturing_enabled:
            warning_message = (
                "画面録画とスクリーンショットが両方オフになっています。\n"
                "少なくとも一方を有効にしてください。"
            )
            logger.warning(
                "タスク開始をブロックしました: 録画・スクリーンショットが共に無効です。"
            )
            messagebox.showwarning("処理を開始できません", warning_message)
            self._update_start_button_state()
            self._unlock_settings_tab()  # P2: アボート時に設定タブを復元
            return

        # フォルダ設定は共通で最初に行う
        if not self.prepare_save_folder():
            self._unlock_settings_tab()  # P2: アボート時に設定タブを復元
            return

        tasks_started = False

        if self.capturing_enabled:
            self.start_screenshot_capture()
            tasks_started = True
        else:
            logger.info("スクリーンショットはユーザー設定によりスキップされます。")
            self.screenshot_status_label.config(
                text="待機中... (スクリーンショットは無効化されています)"
            )

        if self.recording_enabled:
            if hwnd != 0:
                self.start_recording(hwnd)  # 引数でハンドルを渡す
                tasks_started = True
            else:
                if self.capturing_enabled:
                    logger.info(
                        "録画対象ウィンドウが選択されていないため、スクリーンショットのみ開始します。"
                    )
                else:
                    logger.warning(
                        "録画が有効ですが録画対象ウィンドウが選択されていないため、録画を開始できません。"
                    )
                    messagebox.showwarning(
                        "録画対象未選択",
                        "録画を開始するにはウィンドウを選択してください。",
                    )
                if not self.is_recording:
                    self.recording_status_label.config(
                        text="待機中... (録画対象ウィンドウが未選択です)"
                    )
        else:
            logger.info("録画はユーザー設定によりスキップされます。")
            if not self.is_recording:
                self.recording_status_label.config(
                    text="待機中... (録画は無効化されています)"
                )

        if not tasks_started:
            self._update_start_button_state()
            return

        # 画面が暗くならないようスリープ防止を有効化
        try:
            self._prevent_sleep()
        except Exception:
            pass

        # 統合ボタンの状態更新
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.folder_entry.config(state=tk.DISABLED)
        self.refresh_window_list_button.config(
            state=tk.DISABLED
        )  # タスク実行中は更新不可
        self.window_listbox.config(state=tk.DISABLED)
        self.model_combobox.config(state=tk.DISABLED)  # モデル選択も無効化
        self._set_checkbox_state(tk.DISABLED)

    # --- 録画関連メソッド (修正) ---
    def start_recording(self, hwnd):  # 引数 hwnd を追加
        """指定されたウィンドウハンドルで録画を開始する"""
        # hwnd = self.selected_window_handle.get() # start_tasks から渡される
        if hwnd == 0:  # ハンドルが無効な場合は何もしない（基本的には呼ばれないはず）
            logger.error("録画開始エラー: 無効なウィンドウハンドルです。")
            return

        if self.is_recording:  # 既に録画中の場合は何もしない
            return

        logger.info(f"録画を開始します。対象ウィンドウハンドル: {hwnd}")

        # 一時ファイル用ディレクトリを準備
        save_folder = self.save_folder_name.get()
        self.temp_frames_dir = os.path.join(save_folder, "_temp_frames")
        os.makedirs(self.temp_frames_dir, exist_ok=True)
        self.temp_audio_file = os.path.join(save_folder, "_temp_audio.raw")
        logger.info(f"一時ファイル保存先: {self.temp_frames_dir}")

        # 録画スレッドを開始
        self.frame_count = 0  # フレームカウンタ
        self.frame_timestamps = []  # タイムスタンプのみメモリに保持
        self.recording_thread = threading.Thread(
            target=self.recording_loop,
            args=(hwnd,),
            name="RecordingThread",
            daemon=True,
        )
        self.is_recording = True  # is_recording は先に True にする
        self.recording_thread.start()

        # 音声録音スレッドを開始
        self.audio_file_handle = None  # 音声ファイルハンドル
        self.audio_recording_thread = threading.Thread(
            target=self.audio_recording_loop, name="AudioRecordingThread", daemon=True
        )
        self.audio_recording_thread.start()

        # self.start_recording_button.config(state=tk.DISABLED) # 個別ボタン削除
        # self.stop_recording_button.config(state=tk.NORMAL) # 個別ボタン削除
        # ボタン状態は start_tasks / stop_all_tasks で制御
        self.recording_start_time = time.time()
        self.last_sound_time = time.time()  # 録画開始時は音があったとみなす
        self.update_recording_status()  # ステータス更新開始

    def recording_loop(self, hwnd):
        """指定されたウィンドウのフレームを1FPSで録画するループ処理（メモリ効率改善版）"""
        logger.info(f"録画ループを開始します。対象ウィンドウハンドル: {hwnd}")

        try:
            # mssのスクリーンショット用オブジェクト
            sct = mss.mss()

            # ウィンドウの位置とサイズを取得
            try:
                window_rect = win32gui.GetWindowRect(hwnd)
                left, top, right, bottom = window_rect
                width = right - left
                height = bottom - top
                logger.info(
                    f"対象ウィンドウの位置とサイズ: ({left}, {top}, {width}, {height})"
                )

                # ウィンドウ領域の定義
                monitor = {"left": left, "top": top, "width": width, "height": height}
            except Exception as e:
                logger.exception(f"ウィンドウ情報の取得に失敗しました: {e}")
                self.error_occurred_in_recording_thread = True
                return

            # 録画開始時刻を記録
            start_time = time.time()
            local_frame_count = 0

            while self.is_recording:
                # --- ウィンドウ存在チェック ---
                if not win32gui.IsWindow(hwnd):
                    logger.warning(
                        f"録画対象ウィンドウ (HWND: {hwnd}) が見つかりません。録画を停止します。"
                    )
                    # UIスレッドで停止処理を呼び出す
                    self.root.after(0, self.stop_all_tasks)
                    break  # ループを抜ける
                # --- ここまで追加 ---

                try:
                    # 現在のフレーム時刻
                    current_time = time.time()
                    elapsed_time = current_time - start_time

                    # スクリーンショットを取得
                    screenshot = sct.grab(monitor)

                    # Pillowイメージに変換
                    img = Image.frombytes("RGB", screenshot.size, screenshot.rgb)

                    # OpenCV形式に変換
                    frame = np.array(img)
                    frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)

                    # フレームをディスクに保存（メモリに保持しない）
                    if frame is not None and frame.size > 0:
                        frame_filename = os.path.join(
                            self.temp_frames_dir, f"frame_{local_frame_count:06d}.png"
                        )
                        # PNG形式でエンコードして保存（日本語パス対応）
                        success, buffer = cv2.imencode(".png", frame)
                        if success:
                            with open(frame_filename, "wb") as f:
                                f.write(buffer)
                            self.frame_timestamps.append(elapsed_time)
                            local_frame_count += 1
                            self.frame_count = local_frame_count

                            if local_frame_count % 10 == 0:  # 10フレームごとにログ出力
                                logger.info(
                                    f"録画フレーム数: {local_frame_count}, 経過時間: {elapsed_time:.2f}秒"
                                )
                        else:
                            logger.warning("フレームのエンコードに失敗しました")
                    else:
                        logger.warning("無効なフレームがスキップされました")

                    # 次のフレームタイミングまで待機（1FPS）
                    next_frame_time = start_time + (local_frame_count * 1.0)  # 1秒間隔
                    sleep_time = max(0, next_frame_time - time.time())
                    if sleep_time > 0:
                        time.sleep(sleep_time)

                except Exception as e:
                    logger.exception(f"フレーム取得中にエラーが発生しました: {e}")
                    self.error_occurred_in_recording_thread = True
                    time.sleep(1.0)  # エラー時は少し待機してリトライ

        except Exception as e:
            logger.exception(f"録画ループでエラーが発生しました: {e}")
            self.error_occurred_in_recording_thread = True

        finally:
            logger.info(f"録画ループを終了します。合計フレーム数: {self.frame_count}")

    def audio_recording_loop(self):
        """SoundCardを使用してシステムサウンド (ループバック) を録音するループ（メモリ効率改善版）"""
        logger.info("音声録音ループ (SoundCard) を開始します")
        num_frames = 1024  # 一度に読み取るサンプル数 (SoundCardの推奨値に合わせる)

        try:
            # デフォルトのループバックデバイスを取得
            # SoundCard はデフォルトでスピーカーのループバックを選択しようとします
            microphone_device = sc.get_microphone(
                id=str(sc.default_speaker().name), include_loopback=True
            )
            logger.info(
                f"取得したマイクデバイス: {microphone_device.name}"
            )  # マイクデバイス名を確認
            logger.info(f"マイクデバイスの型: {type(microphone_device)}")

            # サンプルレートとチャンネル数をマイクデバイスから取得
            try:
                self.audio_sample_rate = 48000  # デフォルト値
                self.audio_channels = 2  # デフォルト値
                logger.info(
                    f"デバイス情報 - サンプルレート: {self.audio_sample_rate}, チャンネル数: {self.audio_channels}"
                )
            except AttributeError as e:
                logger.error(
                    f"マイクデバイスからサンプルレートまたはチャンネル数を取得できませんでした: {e}"
                )
                # デフォルト値を設定するか、エラー処理を行う

                logger.warning(
                    f"デフォルト値を使用します - サンプルレート: {self.audio_sample_rate}, チャンネル数: {self.audio_channels}"
                )

            # 音声データを一時ファイルに書き出す
            self.audio_file_handle = open(self.temp_audio_file, "wb")
            logger.info(f"音声データ一時ファイル: {self.temp_audio_file}")

            with microphone_device.recorder(
                samplerate=self.audio_sample_rate,  # デバイスから取得した値を使用
                channels=self.audio_channels,  # デバイスから取得した値を使用
                blocksize=num_frames,
            ) as mic:
                logger.info(
                    f"レコーダーオブジェクトの型: {type(mic)}"
                )  # mic の型を確認
                logger.info(
                    f"録音中: {microphone_device.name} (ループバック)"
                )  # マイクデバイス名を使用

                # 録音ループ
                while self.is_recording:
                    try:
                        # recordメソッドで指定したサンプル数を読み取る
                        data = mic.record(numframes=num_frames)
                        if data is not None and data.size > 0:
                            # --- 無音状態チェック ---
                            max_amplitude = np.max(np.abs(data))
                            current_time = time.time()
                            if max_amplitude < self.silence_threshold:
                                # 無音状態
                                if self.last_sound_time is not None:
                                    no_sound_duration = (
                                        current_time - self.last_sound_time
                                    )
                                    if (
                                        no_sound_duration
                                        > self.no_sound_timeout_seconds
                                    ):
                                        logger.info(
                                            f"{self.no_sound_timeout_seconds} 秒以上無音状態が続いたため、録画を停止します。"
                                        )
                                        self.root.after(0, self.stop_all_tasks)
                                        break  # ループを抜ける
                            else:
                                # 音声あり
                                self.last_sound_time = current_time
                            # --- ここまで追加 ---

                            # SoundCard は float32 の NumPy 配列を返す
                            # メモリに保持せず、直接ファイルに書き出す
                            if self.audio_file_handle:
                                data.tofile(self.audio_file_handle)
                        else:
                            logger.warning(
                                "SoundCard から None または空のデータが返されました。"
                            )
                            time.sleep(0.01)  # 少し待機
                    except Exception as e:
                        logger.exception(
                            f"音声データ読み取り/ファイル書き出し中にエラー: {e}"
                        )
                        self.error_occurred_in_recording_thread = True
                        break  # ループ中断

        except Exception as e:
            logger.exception(f"音声録音 (SoundCard) 中にエラーが発生しました: {e}")
            messagebox.showerror(
                "録音エラー",
                f"音声録音の初期化または実行中にエラーが発生しました:\n{e}",
            )
            self.error_occurred_in_recording_thread = True

        finally:
            # 音声ファイルを閉じる
            if self.audio_file_handle:
                try:
                    self.audio_file_handle.close()
                    logger.info("音声一時ファイルを閉じました")
                except Exception as e:
                    logger.error(f"音声ファイルのクローズに失敗: {e}")
            logger.info("音声録音ループ (SoundCard) を終了します")

    def _save_video_with_audio(self, output_filepath):
        """録画されたフレームと音声データからMoviePyを使って動画ファイルを生成・保存し、成功したらノート作成をトリガーする（メモリ効率改善版）"""
        if self.frame_count == 0:
            logger.error("保存するフレームがありません")
            return False

        video_clip = None  # 初期化
        audio_clip = None  # 初期化
        final_clip = None  # 初期化

        try:
            # 保存先のディレクトリを確認
            output_dir = os.path.dirname(output_filepath)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)

            # フレーム画像のパスリストを作成
            frame_files = []
            for i in range(self.frame_count):
                frame_filename = os.path.join(
                    self.temp_frames_dir, f"frame_{i:06d}.png"
                )
                if os.path.exists(frame_filename):
                    frame_files.append(frame_filename)
                else:
                    logger.warning(
                        f"フレームファイルが見つかりません: {frame_filename}"
                    )

            if not frame_files:
                logger.error(
                    "有効なフレームファイルがありませんでした。動画を保存できません。"
                )
                return False

            logger.info(f"合計 {len(frame_files)} フレームを動画クリップに使用します。")

            # フレーム間の時間を計算してFPSを推定
            if len(self.frame_timestamps) > 1:
                avg_interval = np.mean(np.diff(self.frame_timestamps))
                fps = 1.0 / avg_interval if avg_interval > 0 else 1.0
                logger.info(f"推定FPS: {fps:.2f}")
            else:
                fps = 1.0  # フレームが1つしかない場合
                logger.warning("フレームが1つしかないため、FPSを1.0に設定します。")

            # MoviePyのVideoClipを作成（ファイルから直接読み込み）
            video_clip = mpe.ImageSequenceClip(frame_files, fps=fps)

            # 音声データがある場合、AudioClipを作成して結合
            audio_clip = None
            final_clip = video_clip  # デフォルトは音声なし

            # 音声ファイルが存在するか確認
            if (
                os.path.exists(self.temp_audio_file)
                and os.path.getsize(self.temp_audio_file) > 0
            ):
                try:
                    logger.info(f"音声ファイルを読み込みます: {self.temp_audio_file}")
                    # 音声データをファイルから読み込む
                    audio_array = np.fromfile(self.temp_audio_file, dtype=np.float32)

                    if audio_array.size > 0 and self.audio_sample_rate is not None:
                        logger.info(
                            f"音声データを読み込みました。サンプル数: {len(audio_array)}, dtype: {audio_array.dtype}"
                        )

                        # 音声配列の形状を確認・整形
                        if audio_array.ndim == 1 and self.audio_channels == 2:
                            # モノラルデータが連結されて1次元になっている場合、2チャンネルに整形
                            logger.info("1次元の音声配列を2チャンネルに整形します。")
                            try:
                                audio_array = audio_array.reshape(
                                    -1, self.audio_channels
                                )
                            except ValueError as reshape_err:
                                logger.error(
                                    f"音声配列の整形に失敗しました: {reshape_err}。音声なしで続行します。"
                                )
                                audio_array = None
                        elif audio_array.ndim == 1 and self.audio_channels == 1:
                            # 1チャンネルの場合はそのままで良いことが多い
                            pass
                        elif (
                            audio_array.ndim == 2
                            and audio_array.shape[1] == self.audio_channels
                        ):
                            # 既に正しい形式
                            pass
                        else:
                            logger.error(
                                f"音声配列の形状 ({audio_array.shape}) がチャンネル数 ({self.audio_channels}) と一致しません。音声なしで続行します。"
                            )
                            audio_array = None  # 不正な場合は音声なしに

                        if audio_array is not None:
                            try:
                                # AudioArrayClip を直接使用
                                audio_clip = AudioArrayClip(
                                    audio_array, fps=self.audio_sample_rate
                                )
                                # 動画の長さに合わせて音声クリップの長さを調整
                                if audio_clip.duration > video_clip.duration:
                                    audio_clip = audio_clip.subclip(
                                        0, video_clip.duration
                                    )
                                elif audio_clip.duration < video_clip.duration:
                                    # 必要に応じて無音を追加するか、動画を短くする
                                    logger.warning(
                                        f"音声クリップ ({audio_clip.duration:.2f}s) が動画クリップ ({video_clip.duration:.2f}s) より短いです。"
                                    )

                                logger.info(
                                    f"AudioArrayClipを作成しました。Duration: {audio_clip.duration:.2f}秒"
                                )
                                final_clip = video_clip.set_audio(audio_clip)
                                logger.info("動画と音声を結合しました。")
                            except Exception as audio_clip_err:
                                logger.exception(
                                    f"AudioArrayClipの作成または結合中にエラー: {audio_clip_err}"
                                )
                                final_clip = video_clip  # エラー時は音声なし
                        else:
                            final_clip = (
                                video_clip  # 整形失敗などで音声なしになった場合
                            )
                            logger.info("音声データが不正なため、動画のみ保存します。")
                    else:
                        logger.info("音声データが空のため、動画のみ保存します。")

                except Exception as e:
                    logger.exception(f"音声処理中に予期せぬエラーが発生しました: {e}")
                    final_clip = video_clip  # エラー時は音声なしで保存
            else:
                logger.info("音声データファイルが存在しないため、動画のみ保存します。")

            # 動画ファイルを書き出し
            logger.info(f"動画ファイルを書き出します: {output_filepath}")
            try:
                # codec='libx264' を指定して互換性を高める
                # audio_codec='aac' を指定 (多くのプレイヤーでサポート)
                # threads を指定してエンコードを高速化 (CPUコア数など)
                # preset='medium' などで品質と速度のバランスを取る
                final_clip.write_videofile(
                    output_filepath,
                    codec="libx264",
                    audio_codec="aac",
                    temp_audiofile="temp-audio.m4a",  # 一時ファイル名を指定
                    remove_temp=True,  # 一時ファイルを削除
                    threads=8,  # CPUコア数に合わせて調整
                    preset="medium",  # 品質と速度のバランス
                    logger=None,
                    ffmpeg_params=[
                        "-profile:v",
                        "baseline",  # H.264 Baseline Profile: 幅広い互換性
                        "-level",
                        "3.0",  # レベル3.0: 多くのデバイスでサポート
                        "-pix_fmt",
                        "yuv420p",  # ピクセルフォーマット: 最も一般的な形式
                        "-vf",
                        "scale=trunc(iw/2)*2:trunc(ih/2)*2",  # 解像度を偶数に調整 (互換性向上)
                    ],  # MoviePyのログを無効化 (Pythonのloggingを使用するため)
                )
                logger.info(f"動画ファイルを保存しました: {output_filepath}")
                # self.last_saved_video_filename = output_filepath # ここでは設定しない

                # ★★★ 保存成功後にノート作成をトリガー ★★★
                # UIスレッド経由で start_note_creation を呼び出す
                logger.info(
                    f"動画保存成功、ノート作成をトリガーします: {output_filepath}"
                )
                self.root.after(0, self.start_note_creation, output_filepath)
                return True  # 保存自体は成功
            except subprocess.CalledProcessError as e:
                # ffmpeg のエラーメッセージを取得しようとする試み
                ffmpeg_error = ""
                if hasattr(e, "stderr"):

                    try:
                        ffmpeg_error = e.stderr.decode("utf-8", errors="ignore")
                        logger.error(f"FFmpegエラー出力:\n{ffmpeg_error}")
                    except Exception as decode_err:
                        logger.error(f"FFmpegエラー出力のデコードに失敗: {decode_err}")
                logger.exception(
                    f"MoviePyによる動画書き出し中にエラーが発生しました: {e}"
                )
                messagebox.showerror(
                    "動画保存エラー",
                    f"動画の保存に失敗しました:\n{e}\n\nFFmpegエラー:\n{ffmpeg_error[:500]}...",
                )  # エラーメッセージを短縮
                return False

        except Exception as e:
            logger.exception(f"動画保存処理全体でエラーが発生しました: {e}")
            messagebox.showerror(
                "動画保存エラー", f"動画の保存中に予期せぬエラーが発生しました:\n{e}"
            )
            return False
        finally:
            # クリップオブジェクトを閉じる (メモリ解放)
            if "video_clip" in locals() and video_clip:
                video_clip.close()
            if "audio_clip" in locals() and audio_clip:
                audio_clip.close()
            if "final_clip" in locals() and final_clip:
                final_clip.close()

            # 一時ファイルをクリーンアップ
            try:
                if hasattr(self, "temp_frames_dir") and os.path.exists(
                    self.temp_frames_dir
                ):
                    import shutil

                    shutil.rmtree(self.temp_frames_dir, ignore_errors=True)
                    logger.info(
                        f"一時フレームディレクトリを削除しました: {self.temp_frames_dir}"
                    )
            except Exception as cleanup_err:
                logger.error(f"一時フレームディレクトリの削除に失敗: {cleanup_err}")

            try:
                if hasattr(self, "temp_audio_file") and os.path.exists(
                    self.temp_audio_file
                ):
                    os.remove(self.temp_audio_file)
                    logger.info(
                        f"一時音声ファイルを削除しました: {self.temp_audio_file}"
                    )
            except Exception as cleanup_err:
                logger.error(f"一時音声ファイルの削除に失敗: {cleanup_err}")

            logger.info("動画保存処理を終了します。")

    def stop_recording(self):
        """録画と音声録音を停止し、動画ファイルを保存する"""
        if not self.is_recording:
            return

        logger.info("録画停止処理を開始します。")
        self.is_recording = False  # まずフラグを False にする

        # 録画スレッドの終了を待つ
        if self.recording_thread and self.recording_thread.is_alive():
            logger.info("録画スレッドの終了を待機します...")
            self.recording_thread.join(timeout=5.0)  # タイムアウトを設定
            if self.recording_thread.is_alive():
                logger.warning("録画スレッドが時間内に終了しませんでした。")
            else:
                logger.info("録画スレッドが終了しました。")
        self.recording_thread = None

        # 音声録音スレッドの終了を待つ
        if self.audio_recording_thread and self.audio_recording_thread.is_alive():
            logger.info("音声録音スレッドの終了を待機します...")
            self.audio_recording_thread.join(timeout=5.0)  # タイムアウトを設定
            if self.audio_recording_thread.is_alive():
                logger.warning("音声録音スレッドが時間内に終了しませんでした。")
            else:
                logger.info("音声録音スレッドが終了しました。")
        self.audio_recording_thread = None

        # ボタン状態は stop_all_tasks で制御

        # 動画ファイルを保存
        if self.frame_count > 0:
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"recording_{now}.mp4"
            # ★★★ 保存パスは self.save_folder_name.get() から取得 ★★★
            save_folder = self.save_folder_name.get()
            if not save_folder or not os.path.isdir(save_folder):
                logger.error(
                    f"無効な保存フォルダです: {save_folder}。動画を保存できません。"
                )
                # 必要であればここでエラーメッセージ表示
                messagebox.showerror(
                    "保存エラー",
                    f"指定された保存フォルダが見つかりません:\n{save_folder}",
                )
            else:
                output_filepath = os.path.join(save_folder, output_filename)
                logger.info(f"動画ファイルの保存を開始します: {output_filepath}")

                # 保存処理を別スレッドで行う (UIが固まるのを防ぐため)
                save_thread = threading.Thread(
                    target=self._save_video_with_audio_background,
                    args=(output_filepath,),
                    name="VideoSaveThread",
                    daemon=True,
                )
                save_thread.start()
                logger.info("動画保存スレッドを開始しました。")

        else:
            logger.warning("録画フレームがないため、動画ファイルは保存されません。")

        self.recording_start_time = None
        self.update_recording_status()  # ステータスを更新

        # エラーチェック
        if self.error_occurred_in_recording_thread:
            messagebox.showerror(
                "録画エラー",
                "録画中にエラーが発生しました。詳細はログを確認してください。",
            )
            self.error_occurred_in_recording_thread = False  # フラグをリセット

        logger.info("録画停止処理を完了しました。")

    def _restore_gui_after_note_creation(self):
        """ノート作成後または中断時にGUI操作制限を解除する共通処理"""
        logger.info("GUI操作制限を解除します。")
        self._update_start_button_state()
        self.stop_button.config(state=tk.DISABLED)  # 停止ボタンは常に無効で良い
        self.folder_entry.config(state=tk.NORMAL)
        self.refresh_window_list_button.config(state=tk.NORMAL)
        self.window_listbox.config(state=tk.NORMAL)
        self.model_combobox.config(state="readonly")
        self._set_checkbox_state(tk.NORMAL)
        self.root.update()
        # 閉じるボタンを元に戻す
        if hasattr(self, "original_on_closing"):
            self.root.protocol("WM_DELETE_WINDOW", self.original_on_closing)
        else:
            self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
        # スリープ設定を元に戻す
        try:
            self._restore_sleep()
        except Exception:
            pass
        # ノート作成完了後に設定タブを再度有効化
        self._unlock_settings_tab()

    def _prevent_sleep(self):
        """Windowsのスリープ/ディスプレイオフを一時的に防止する。"""
        try:
            # ES_CONTINUOUS | ES_DISPLAY_REQUIRED | ES_SYSTEM_REQUIRED
            ES_CONTINUOUS = 0x80000000
            ES_SYSTEM_REQUIRED = 0x00000001
            ES_DISPLAY_REQUIRED = 0x00000002
            windll.kernel32.SetThreadExecutionState(
                ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
            )
            logger.info("画面スリープ防止を有効化しました。")
        except Exception as e:
            logger.exception(f"画面スリープ防止の設定に失敗しました: {e}")

    def _restore_sleep(self):
        """スリープ制御を元に戻す。"""
        try:
            ES_CONTINUOUS = 0x80000000
            windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
            logger.info("画面スリープ防止を無効化しました。")
        except Exception as e:
            logger.exception(f"画面スリープ制御の解除に失敗しました: {e}")

    def _save_video_with_audio_background(self, output_filepath):
        """動画保存の成否に応じてUI復元処理を行うラッパー"""
        success = self._save_video_with_audio(output_filepath)
        if not success:
            self.root.after(0, self._handle_video_save_failure)

    def _handle_video_save_failure(self):
        """動画保存に失敗した際のUI復元処理"""
        logger.error("動画保存失敗のためUIを復元します。")
        self.recording_status_label.config(text="動画の保存に失敗しました。")
        self._restore_gui_after_note_creation()

    def start_note_creation(self, video_filepath):  # 引数に video_filepath を追加
        """指定された動画ファイルパスでノート作成処理を開始する"""
        # --- GUI操作制限 ---
        logger.info("ノート作成開始: GUI操作を制限します。")
        # ノート作成は stop_all_tasks 後に呼ばれるため、ボタンは既に無効化されているはずだが念のため
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)  # stop_all_tasks で無効化される
        self.refresh_window_list_button.config(state=tk.DISABLED)
        self.window_listbox.config(state=tk.DISABLED)
        # 閉じるボタンを無効化
        self.root.protocol(
            "WM_DELETE_WINDOW",
            lambda: logger.warning("ノート作成中はウィンドウを閉じられません。"),
        )
        # 画面が暗くならないようスリープ防止を有効化
        try:
            self._prevent_sleep()
        except Exception:
            pass
        # --- ここまで ---

        # バリデーションチェック: APIクライアントの確認
        if not self.gemini_client:
            self.note_creation_status.set("ノート作成不可: APIクライアント未設定")
            logger.warning(
                "Gemini API クライアントが設定されていないため、ノート作成を開始できません。"
            )
            # GUI操作制限を解除してから return
            self._restore_gui_after_note_creation()
            return

        # バリデーションチェック: 動画ファイルの存在確認
        if not video_filepath or not os.path.exists(video_filepath):
            self.note_creation_status.set(
                f"ノート作成不可: 動画ファイルが見つかりません ({video_filepath})"
            )
            logger.error(f"指定された動画ファイルが見つかりません: {video_filepath}")
            # GUI操作制限を解除してから return
            self._restore_gui_after_note_creation()
            return

        # ★ 開始時のステータスを設定
        self.note_creation_status.set("ノート生成準備中...")
        self.note_result_message.set("")
        logger.info(f"ノート作成処理を開始します。対象動画: {video_filepath}")

        # Gemini APIの呼び出しとmd生成の処理 (内部関数)
        def api_call_and_md_gen():
            video_file = None
            try:
                prompt = """
# 目的

添付された動画の内容を分析し、指定されたJSONスキーマに従って、各トピックごとにノート作成に必要な情報を抽出せよ。

## 基本ルール

- 想定読者は「その分野の専門外の大学生」。専門的な表現を避け、平易な日本語で記述する。
- 常に日本語で回答する。入力が他の言語であっても翻訳せず日本語で回答する。
- 動画内の主題に無関係な内容（雑談、環境音など）は除外する。
- 動画内容を慎重に分析し、漏れや誤りがないことを確認する。
- ~です。~ます。調は使わない。
- 思考は英語で行い、出力は日本語で行う。
- 時間がどれだけかかろうと、推論は限界まで行うこと（推論コストの節約は悪である）。

# トピック分割基準

動画内容をトピックに分割する際、以下のルールに従う。

## 分割の原則

**「話題の転換点」** でトピックを分割する。

話題の転換とは、講義内の **主題、対象、目的** のいずれかが切り替わるタイミングである。
同一の主題内でも目的が変わる場合（例:「基礎知識の説明→実験の紹介→結果の考察」）は別トピックとして扱う。

## 転換のシグナル例

以下のいずれかに該当する場合、新しいトピックの開始と判断する：

- 講師が明示的に話題を切り替えた（「次に」「ここから」「本題に入ります」等）
- 説明の対象となる概念や主題が変わった（例:「分子の性質」→「実験手順」）
- 議論の目的が変わった（例:「メカニズムの説明」→「結果から示唆される議論」）
- スライドや図の切り替えにより、明らかに異なる論点に移行した

## 必ず独立トピックとすべき要素

- **導入/概要**: 講義冒頭で全体の目的や背景を説明している部分
- **考察/まとめ**: 実験結果から導かれる知見や意義を議論している部分で、独立した内容がある場合
- **課題/宿題の説明**: 講義本編とは異なる情報のため混在させない（assignmentsフィールドに別途記録）

## トピック粒度の目安

- 1トピックの目安は **1〜3分** 程度。4分を超える場合は、内部に複数の論点が含まれていないか再確認し、分割を検討する。
- 一つのトピック内に **異なる目的の議論**（例:「実験手法の説明」と「結果の考察」と「結論」）が混在している場合は、必ず分割する。
- 講義全体で **5トピック以上** になることが一般的である。極端に少ないトピック数（3〜4個）は、分割が粗すぎる可能性が高い。

# topic_summary（トピック要約）の記述基準

topic_summaryは、そのトピックで何が話されたかが読むだけで分かる要約でなければならない。
**「〜について説明している」のようなメタ的紹介文は禁止。**
以下の3要素を含むこと：

1. **何が話されたか**（主題・対象）
2. **どのような論点・主張があったか**（議論の方向性）
3. **どのような結論・知見が示されたか**（帰結・結果）

## 良い例・悪い例

❌ 悪い例:「両親媒性分子の定義と、それらが水中で形成する構造について説明している。」
→ 何が説明されたかは分かるが、内容そのものが全く伝わらない。

✅ 良い例:「両親媒性分子は親水基と疎水基を持ち、水中では濃度に応じてミセル（単層球状構造）やリポソーム（二重層球状構造）を自発的に形成する。特にリポソームは内部に水相を保持できるため、人工細胞の容器モデルとして適している。」
→ 具体的な内容と結論が一文で把握できる。

# topic_points（重要ポイント）の記述基準

topic_pointsは、"暗記"ではなく"理解の支点"となるポイントを箇条書きで記述する。
以下のルールに従うこと：

- **具体的な数値・条件を含める**: 講義中で言及された測定値、温度、距離、収率、濃度、速度などの定量的データは省略せず記載する。
  - ❌ 悪い例:「高い電荷移動度を示す」
  - ✅ 良い例:「0.02 cm²V⁻¹s⁻¹の電荷移動度を示す」
  - ❌ 悪い例:「分子間距離はファンデルワールス半径に近い」
  - ✅ 良い例:「分子間距離は3.39 Åで、C–Cファンデルワールス半径の和（3.40 Å）とほぼ一致する」
- **未解明点・限界・今後の課題**: 講義内で「まだ分かっていない」「今後の検討が必要」と言及された事項は、必ず1項目として含める。
  - ✅ 例:「四量体での化学シフト反転の原因は未解明であり、溶媒中の水の影響が一因として議論されている」
  - ✅ 例:「非会合型OHへのシフトは今回の測定範囲では観測されておらず、より詳細な条件での検証が必要」

# important_knowledge（重要な知見）の記述基準

important_knowledgeは、そのトピックで最も重要な知識を構造化して記録するフィールドである。
以下のルールに従うこと：

## statement の基準

- 1〜2文で、知識の核心を言い切る形で記述する。
- 講義中に提示された具体的な数値（測定値、距離、温度、濃度など）がある場合は、statementに含める。

## explanation の基準

- 「なぜ重要か」「どういう条件で成り立つか」を補足する。
- 実験結果に基づく場合は、使用された分析手法と結果の概要を含める。

## 未解明点・限界の扱い

- 講義内で「未解明」「不明」「今後の課題」として言及された事項がある場合、それも1つのimportant_knowledgeとして記録する。
  - source は `lecture_explicit` とする。
  - statementに「〜は未解明である」「〜の検証が今後必要とされている」等の形で記述する。
  - explanationに、未解明である事情や今後の方針を記述する。

# 専門用語の抽出基準

各トピックの `technical_term` を以下の基準に従って抽出する。各トピックの technical_term は **最低4語** を目標に抽出すること。
ただし、各用語の `explanation` は**必ず以下のexplanation の品質基準セクション**を満たすこと。短い説明文は**絶対禁止**とする。

## 抽出プロセス（2段階で行う）

**Step 1 — 列挙**: トピック内に登場する専門用語を、まず用語名だけ全て洗い出す。この段階では数を絞らず、網羅性を最優先する。

**Step 2 — 詳述**: Step 1で洗い出した全用語について `technical_term` の各フィールドを記述する。Step 1で挙げた用語を省略してはならない。

## 判断基準

- そのトピックを理解するために必要な専門用語を抽出する。
- 抽出の基準は **日常会話では伝わらない語** であること。
- 迷った場合は **含める** 。漏れは後から追加できないが、不要なものは後から除去できる。
- 一度だけ、あるいは一瞬だけ言及された用語も含める。スライド上に数秒表示された試薬名や、口頭で軽く触れた手法名であっても、学生が調べる必要がある語は対象とする。

## 対象カテゴリ（以下に該当するものを積極的に抽出する）

- 学術・科学用語（例: エントロピー、ホメオスタシス）
- 業界固有の略語・頭字語（例: ROI、PDCA、API）
- 法律・制度用語（例: 品質保証、善意の第三者）
- 数学・統計の記法や用語（例: 標準偏差、回帰分析、偏微分）
- 講義固有の定義を持つ用語（講師が独自の意味や文脈で使用する語）
- 意味が曖昧なカタカナ外来語（例: アジャイル、レジリエンス）

## explanation の品質基準

`technical_term` の各 `explanation` は以下の品質基準を **厳守** すること：

- **長さ**: **最低3文**。1〜2文だけの辞書的定義は **絶対禁止** 。上限は5文程度。
- **内容**: 必ず以下の **3点すべて** を含む：
  1. その概念が何であるか（定義）
  2. このトピックの文脈でなぜ重要か、どう使われているか（文脈）
  3. 理解を助ける補足情報（以下のいずれか1つ以上）：
     - 類似概念・対比概念との違い
     - 具体的な数値・条件・適用例
     - 実用上の意義や応用先
     - 身近な比喩やイメージ

### 悪い例（不合格）

- ❌「化学反応を促進する物質。」→ 1文、文脈なし、補足なし。
- ❌「外部からの制御なしに、分子自らが自発的に秩序ある構造を形成するプロセス。」→ 定義のみで文脈と補足がない。
- ❌「合成した物質の構造や性質を、様々な分析手法を用いて明らかにすること。」→ 1文の辞書定義。

### 良い例（合格）

- ✅「化学反応を促進するが、それ自体は消費されない物質。このトピックでは、原料から機能的な構成要素への変換において中核的な役割を果たし、議論されている自己持続サイクルを可能にしている。工業的には反応効率を大幅に改善するために不可欠な存在である。」
- ✅「大きな環状構造を持つ分子の総称。内部に空孔を持つため、特定の小さな分子やイオンを「ゲスト」として取り込むホスト・ゲスト化学において中心的な役割を果たす。本トピックでは、水分子を包接するナノチューブの構成単位として用いられている。」

# 備考

- 具体的なJSONスキーマの構造は別途提供される前提で処理する。
- topicsセクションでは、推論の最終結論として整理された出力を記述する。
- technical_term と important_knowledge は単一とは限らず、複数存在しうる。

# タイムスタンプルール

- タイムスタンプは「動画再生位置（動画冒頭からの経過時間＝タイムコード）」を表す。
- ローカル時間やUTCなどの「時刻」ではない。日付やタイムゾーン（+09:00、Zなど）は含めない。
- フォーマットは MM:SS（ゼロ埋め、固定2桁）。例: 03:31
- 常に start_time < end_time を満たすこと。
- evidence は topic_time_range 内に収めること。
- 不確実な場合は推測せず、evidence: [] とする。
- スキーマに基づく適切な階層関係を確保する（例: トピックのタイムスタンプ範囲 > important_knowledgeのタイムスタンプ範囲）。

"""
                # スキーマ定義は変更なし (省略)
                schema = {
                    "type": "object",
                    "properties": {
                        "title": {
                            "type": "string",
                            "description": "動画（講義）のタイトル",
                        },
                        "summary": {
                            "type": "string",
                            "description": "動画全体の要約（目安: 3000字以内）。講義の主張・流れ・結論が分かる形で。",
                        },
                        "topics": {
                            "type": "array",
                            "description": "動画内のトピック",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "topic_title": {
                                        "type": "string",
                                        "description": "トピックのタイトル",
                                    },
                                    "topic_time_range": {
                                        "type": "object",
                                        "description": "そのトピックに関するタイムスタンプ（取れる場合のみ）。",
                                        "properties": {
                                            "start_time": {
                                                "type": "string",
                                                "description": "動画でのトピックの開始時間（フォーマット:MM:SS）。例: 12:34",
                                            },
                                            "end_time": {
                                                "type": "string",
                                                "description": "動画でのトピックの終了時間（フォーマット:MM:SS）。例: 18:02",
                                            },
                                        },
                                        "required": ["start_time", "end_time"],
                                        "propertyOrdering": ["start_time", "end_time"],
                                    },
                                    "topic_keywords": {
                                        "type": "array",
                                        "description": "検索用キーワード（任意）。無ければ空配列。",
                                        "items": {"type": "string"},
                                    },
                                    "topic_summary": {
                                        "type": "string",
                                        "description": "トピック要約（目安: 1000字以内）。「〜について説明しています」のようなメタ的紹介文ではなく、論点・主張・結論を含んだ具体的な内容の要約を記述する。",
                                    },
                                    "topic_points": {
                                        "type": "array",
                                        "description": "重要ポイント（箇条書き）。“暗記”ではなく“理解の支点”になる粒度で。",
                                        "items": {"type": "string"},
                                    },
                                    "important_knowledge": {
                                        "type": "array",
                                        "description": "このトピックで“重要だから残す知識”。初出かどうかは厳密に判定せず、重要性優先。含める内容の例: 定義/概念、前提条件、手順/アルゴリズム、式/関係、例（適用ケース）、落とし穴（誤解しやすい点）。",
                                        "items": {
                                            "type": "object",
                                            "properties": {
                                                "statement": {
                                                    "type": "string",
                                                    "description": "知識の核（1〜2文で言い切り）",
                                                },
                                                "explanation": {
                                                    "type": "string",
                                                    "description": "大学生向けの補足説明（なぜ重要か、成立条件、使いどころ等）。長文化しない。",
                                                },
                                                "source": {
                                                    "type": "string",
                                                    "description": "出所: lecture_explicit=講義で明示 / lecture_inferred=講義内容から推論 / model_background=一般知識補完（講義で言っていない可能性あり）",
                                                    "enum": [
                                                        "lecture_explicit",
                                                        "lecture_inferred",
                                                        "model_background",
                                                    ],
                                                },
                                                "evidence": {
                                                    "type": "object",
                                                    "description": "動画内で根拠が触れられている部分のタイムスタンプ。確実でない場合は空配列（捏造を避ける）。",
                                                    "properties": {
                                                        "start_time": {
                                                            "type": "string",
                                                            "description": "根拠の開始時間（MM:SS）",
                                                        },
                                                        "end_time": {
                                                            "type": "string",
                                                            "description": "根拠の終了時間（MM:SS）",
                                                        },
                                                    },
                                                    "required": [
                                                        "start_time",
                                                        "end_time",
                                                    ],
                                                    "propertyOrdering": [
                                                        "start_time",
                                                        "end_time",
                                                    ],
                                                },
                                                "importance": {
                                                    "type": "integer",
                                                    "description": "任意：重要度（1〜5）",
                                                    "minimum": 1,
                                                    "maximum": 5,
                                                },
                                            },
                                            "required": [
                                                "statement",
                                                "explanation",
                                                "source",
                                                "evidence",
                                            ],
                                            "propertyOrdering": [
                                                "statement",
                                                "explanation",
                                                "source",
                                                "evidence",
                                                "importance",
                                            ],
                                        },
                                    },
                                    "technical_term": {
                                        "type": "array",
                                        "description": "トピックで触れられた専門用語とその解説",
                                        "items": {
                                            "type": "object",
                                            "properties": {
                                                "word": {
                                                    "type": "string",
                                                    "description": "専門用語",
                                                },
                                                "explanation": {
                                                    "type": "string",
                                                    "description": "大学生向けの平易な説明（1〜5文）。何のための概念か/どこで使うかが分かる形。",
                                                },
                                                "lecture_definition": {
                                                    "type": "string",
                                                    "description": "講義内で明示された定義（取れる場合のみ。取れなければ省略）。",
                                                },
                                                "pitfall": {
                                                    "type": "string",
                                                    "description": "誤解・混同ポイント･あると学習が捗る補足説明（あれば。無ければ省略）。",
                                                },
                                                "evidence": {
                                                    "type": "array",
                                                    "description": "用語が説明・登場した根拠タイムスタンプ（任意）。捏造防止のため必須にしない。",
                                                    "items": {
                                                        "type": "object",
                                                        "properties": {
                                                            "start_time": {
                                                                "type": "string",
                                                                "description": "根拠の開始時間（MM:SS）",
                                                            },
                                                            "end_time": {
                                                                "type": "string",
                                                                "description": "根拠の終了時間（MM:SS）",
                                                            },
                                                        },
                                                        "required": [
                                                            "start_time",
                                                            "end_time",
                                                        ],
                                                        "propertyOrdering": [
                                                            "start_time",
                                                            "end_time",
                                                        ],
                                                    },
                                                },
                                                "frequency": {
                                                    "type": "string",
                                                    "description": "この用語がトピック内で登場した頻度。repeated=複数回言及 / once=一度だけ言及 / background=スライド表示のみ等。",
                                                    "enum": [
                                                        "repeated",
                                                        "once",
                                                        "background",
                                                    ],
                                                },
                                            },
                                            "required": ["word", "explanation"],
                                            "propertyOrdering": [
                                                "word",
                                                "explanation",
                                                "lecture_definition",
                                                "pitfall",
                                                "evidence",
                                                "frequency",
                                            ],
                                        },
                                    },
                                },
                                "required": [
                                    "topic_title",
                                    "topic_summary",
                                    "topic_points",
                                    "topic_keywords",
                                    "important_knowledge",
                                ],
                                "propertyOrdering": [
                                    "topic_title",
                                    "topic_time_range",
                                    "topic_keywords",
                                    "topic_summary",
                                    "topic_points",
                                    "important_knowledge",
                                    "technical_term",
                                ],
                            },
                        },
                        "assignments": {
                            "type": "array",
                            "description": "課題・提出物・今後やるべきこと（講義内で指示があったもの中心）。",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "task_title": {
                                        "type": "string",
                                        "description": "課題/タスク名（例: レポート1、演習#3、次回までの予習 など）",
                                    },
                                    "task_detail": {
                                        "type": "string",
                                        "description": "内容・要求（提出物、分量、形式、評価観点、参照範囲など）",
                                    },
                                    "due_date": {
                                        "type": "string",
                                        "format": "date",
                                        "description": "提出期限（日付のみが分かる場合のみ記入）。例: 2026-02-20",
                                    },
                                    "due_datetime": {
                                        "type": "string",
                                        "format": "date-time",
                                        "description": "提出期限（日時まで分かる場合）。例: 2026-02-20T23:59:00+09:00",
                                    },
                                    "evidence": {
                                        "type": "array",
                                        "description": "課題指示が出た根拠となるタイムスタンプ。確実でない場合は空配列でよい。",
                                        "items": {
                                            "type": "object",
                                            "properties": {
                                                "start_time": {
                                                    "type": "string",
                                                    "description": "根拠の開始時間（MM:SS）。例: 42:10",
                                                },
                                                "end_time": {
                                                    "type": "string",
                                                    "description": "根拠の終了時間（MM:SS）。例: 43:05",
                                                },
                                            },
                                            "required": ["start_time", "end_time"],
                                            "propertyOrdering": [
                                                "start_time",
                                                "end_time",
                                            ],
                                        },
                                    },
                                },
                                "required": ["task_title", "task_detail"],
                                "propertyOrdering": [
                                    "task_title",
                                    "task_detail",
                                    "due_date",
                                    "due_datetime",
                                    "evidence",
                                ],
                            },
                        },
                    },
                    "required": ["title", "summary", "topics"],
                    "propertyOrdering": ["title", "summary", "topics", "assignments"],
                }

                logger.info(f"動画ファイルをアップロード開始: {video_filepath}")
                # --- ステータス更新: アップロード中 ---
                self.root.after(
                    0, self.note_creation_status.set, "動画アップロード中..."
                )

                # --- ファイルアップロードと状態待機 (再試行ロジック付き) ---
                max_retries = 3
                retry_delay = 1  # seconds
                video_file = None
                upload_and_wait_success = False

                for attempt in range(max_retries):
                    try:
                        logger.info(
                            f"動画ファイルアップロード試行 {attempt + 1}/{max_retries}: {video_filepath}"
                        )
                        self.root.after(
                            0,
                            self.note_creation_status.set,
                            f"動画ファイルアップロード試行 {attempt + 1}/{max_retries}...",
                        )

                        # 1. ファイルアップロード
                        if self.gemini_client is None:
                            raise Exception(
                                "Gemini API クライアントが設定されていません。"
                            )
                        video_file = self.gemini_client.files.upload(
                            file=video_filepath
                        )
                        logger.info(
                            f"アップロード開始: {video_file.name}, State: {video_file.state}"
                        )
                        self.root.after(
                            0,
                            self.note_creation_status.set,
                            f"アップロード開始: {video_file.state}",
                        )

                        # 2. ファイルがACTIVEになるまで待機
                        polling_interval = 5
                        timeout_seconds = 1800  # 30分
                        start_poll_time = time.time()
                        while video_file.state != "ACTIVE":
                            if time.time() - start_poll_time > timeout_seconds:
                                raise TimeoutError(
                                    f"ファイル処理がタイムアウトしました ({timeout_seconds}秒): {video_file.name}"
                                )

                            # --- ステータス更新: 処理中 ---
                            elapsed_time = time.time() - start_poll_time
                            status_text = f"動画処理中... ({video_file.state}, {elapsed_time:.1f}秒)"
                            self.root.after(
                                0, self.note_creation_status.set, status_text
                            )
                            logger.info(status_text)

                            time.sleep(polling_interval)
                            if video_file.name is None:
                                raise ValueError(
                                    "アップロードされたファイルの名前が取得できません。"
                                )
                            video_file = self.gemini_client.files.get(
                                name=video_file.name
                            )  # 最新の状態を取得

                        logger.info(f"ファイルがACTIVEになりました: {video_file.name}")
                        self.root.after(
                            0,
                            self.note_creation_status.set,
                            "ファイル処理完了 (ACTIVE)",
                        )
                        upload_and_wait_success = True
                        break  # 成功したらループを抜ける

                    except TimeoutError as timeout_err:
                        logger.warning(
                            f"試行 {attempt + 1}/{max_retries}: ファイル処理タイムアウト - {timeout_err}"
                        )
                        # タイムアウトの場合はアップロードされたファイルを削除試行 (失敗しても無視)
                        if video_file and video_file.name:
                            try:
                                logger.info(
                                    f"タイムアウトしたファイル {video_file.name} を削除します。"
                                )
                                if self.gemini_client is None:
                                    raise Exception(
                                        "Gemini API クライアントが設定されていません。"
                                    )
                                self.gemini_client.files.delete(name=video_file.name)
                            except Exception as delete_err:
                                logger.warning(
                                    f"タイムアウトしたファイルの削除中にエラー: {delete_err}"
                                )
                        video_file = None  # ファイル参照をリセット

                        if attempt < max_retries - 1:
                            logger.info(f"{retry_delay}秒後に再試行します...")
                            self.root.after(
                                0,
                                self.note_creation_status.set,
                                f"タイムアウト、再試行 ({attempt + 1}/{max_retries})... {retry_delay}秒待機",
                            )
                            time.sleep(retry_delay)
                            retry_delay *= 2  # 指数バックオフ
                        else:
                            logger.exception(
                                f"ファイル処理が最終的にタイムアウトしました (試行 {max_retries}回): {timeout_err}"
                            )
                            self.root.after(
                                0,
                                self.note_creation_status.set,
                                f"エラー: ファイル処理タイムアウト (試行 {max_retries}回)",
                            )
                            self.root.after(
                                0,
                                self.finish_note_creation,
                                False,
                                f"ファイル処理タイムアウト (試行 {max_retries}回): {timeout_err}",
                            )
                            return  # 処理中断

                    except Exception as upload_err:
                        logger.warning(
                            f"試行 {attempt + 1}/{max_retries}: アップロードまたは状態確認中にエラー - {upload_err}"
                        )
                        # エラー発生時もファイルを削除試行
                        if video_file and video_file.name:
                            try:
                                logger.info(
                                    f"エラーが発生したファイル {video_file.name} を削除します。"
                                )
                                if self.gemini_client is None:
                                    raise Exception(
                                        "Gemini API クライアントが設定されていません。"
                                    )
                                self.gemini_client.files.delete(name=video_file.name)
                            except Exception as delete_err:
                                logger.warning(
                                    f"エラーファイルの削除中にエラー: {delete_err}"
                                )
                        video_file = None  # ファイル参照をリセット

                        if attempt < max_retries - 1:
                            logger.info(f"{retry_delay}秒後に再試行します...")
                            self.root.after(
                                0,
                                self.note_creation_status.set,
                                f"エラー発生、再試行 ({attempt + 1}/{max_retries})... {retry_delay}秒待機",
                            )
                            time.sleep(retry_delay)
                            retry_delay *= 2  # 指数バックオフ
                        else:
                            logger.exception(
                                f"ファイルアップロード/処理が最終的に失敗しました (試行 {max_retries}回): {upload_err}"
                            )
                            self.root.after(
                                0,
                                self.note_creation_status.set,
                                f"エラー: アップロード/処理失敗 (試行 {max_retries}回)",
                            )
                            self.root.after(
                                0,
                                self.finish_note_creation,
                                False,
                                f"アップロード/処理エラー (試行 {max_retries}回): {upload_err}",
                            )
                            return  # 処理中断

                # ループを抜けた後、成功フラグを確認
                if not upload_and_wait_success or video_file is None:
                    logger.error(
                        "不明な理由でファイルアップロード/処理に失敗しました。"
                    )
                    self.root.after(
                        0,
                        self.note_creation_status.set,
                        "エラー: ファイル処理失敗 (不明)",
                    )
                    self.root.after(
                        0, self.finish_note_creation, False, "ファイル処理エラー (不明)"
                    )
                    return

                # --- ステータス更新: ノート生成中 ---
                self.root.after(0, self.note_creation_status.set, "ノート生成中...")

                # Gemini APIに要約リクエストを送信
                logger.info("Gemini APIに要約リクエストを送信します...")
                # ★ モデル名を修正 (例: gemini-1.5-pro-latest)
                #    config ではなく generation_config を使用
                response = self.api_call_with_retry(video_file, prompt, schema)
                if self.gemini_client is None:
                    raise Exception("Gemini API クライアントが設定されていません。")
                if video_file.name is None:
                    raise ValueError(
                        "アップロードされたファイルの名前が取得できません。"
                    )
                self.gemini_client.files.delete(name=video_file.name)
                if response is None or response.text is None:
                    raise ValueError(
                        "Gemini APIの応答が空です。APIの設定やリクエスト内容を確認してください。"
                    )
                summary_text = response.text
                logger.info("Gemini API から応答を取得しました。")
                # --- ステータス更新: 応答処理中 ---
                self.root.after(0, self.note_creation_status.set, "応答を処理中...")

                # ノートファイル生成 (MD or Word)
                try:
                    summary_data = json.loads(summary_text)
                    logger.info("応答のJSONパースに成功しました。")

                    base_name = os.path.splitext(os.path.basename(video_filepath))[0]
                    output_dir = os.path.dirname(video_filepath)
                    os.makedirs(output_dir, exist_ok=True)

                    note_output_mode = getattr(self, "note_output_mode", "MD")

                    if note_output_mode == "Word":
                        # --- Word 出力 ---
                        doc_filename = f"note_{base_name}.docx"
                        doc_filepath = os.path.join(output_dir, doc_filename)
                        logger.info(f"Wordファイルを生成します: {doc_filepath}")
                        doc = Document()
                        populate_word_document(doc, summary_data)
                        doc.save(doc_filepath)
                        logger.info(f"Wordファイル作成完了しました: {doc_filepath}")
                        self.root.after(
                            0, self.finish_note_creation, True, doc_filepath
                        )
                        return
                    else:
                        # --- MD 出力 (デフォルト) ---
                        md_filename = f"note_{base_name}.md"
                        md_filepath = os.path.join(output_dir, md_filename)
                        logger.info(f"Markdownファイルを生成します: {md_filepath}")
                        markdown_content = render_markdown_note(summary_data)
                        with open(md_filepath, "w", encoding="utf-8") as f:
                            f.write(markdown_content)
                        logger.info(f"Markdownファイル作成完了しました: {md_filepath}")
                        self.root.after(0, self.finish_note_creation, True, md_filepath)
                        return

                except json.JSONDecodeError as json_err:
                    logger.error(f"Gemini応答JSON解析失敗: {json_err}")
                    logger.error(f"受信テキスト(一部): {summary_text[:500]}...")
                    error_msg = f"API応答JSON解析エラー: {json_err}\n応答(一部): {summary_text[:200]}..."
                    self.root.after(0, self.finish_note_creation, False, error_msg)
                except Exception as md_err:
                    logger.exception(f"mdファイル生成エラー: {md_err}")
                    self.root.after(
                        0,
                        self.finish_note_creation,
                        False,
                        f"md生成エラー: {md_err}",
                    )

            except TimeoutError as te:
                logger.error(f"ノート作成タイムアウト: {te}")
                self.root.after(
                    0,
                    self.finish_note_creation,
                    False,
                    f"ファイル処理タイムアウト: {te}",
                )
            except Exception as e:
                logger.exception("ノート作成処理中エラー")
                # ★ ファイル削除処理を追加
                if "video_file" in locals() and video_file:
                    try:
                        logger.info(
                            f"エラー発生のためアップロードファイル削除: {video_file.name}"
                        )
                        if self.gemini_client is None:
                            raise Exception(
                                "Gemini API クライアントが設定されていません。"
                            )
                        if video_file.name is None:
                            raise ValueError(
                                "アップロードされたファイルの名前が取得できません。"
                            )
                        self.gemini_client.files.delete(name=video_file.name)
                        logger.info("ファイル削除成功")
                    except Exception as delete_err:
                        logger.error(f"アップロードファイル削除失敗: {delete_err}")
                # UIスレッドでステータスを更新 (失敗)
                self.root.after(0, self.finish_note_creation, False, str(e))
            # ★ finally ブロックを削除 (エラーハンドリング内でファイル削除を行うため)

        # 別スレッドで実行
        note_thread = threading.Thread(
            target=api_call_and_md_gen, name="NoteCreationThread", daemon=True
        )
        note_thread.start()

    def api_call_with_retry(
        self, video_file, prompt, schema, max_retries=3, retry_delay=5
    ):
        """リトライロジックを持つAPIコール"""
        for attempt in range(max_retries):
            try:
                logger.info(f"Gemini API呼び出し試行 {attempt+1}/{max_retries}")
                # 選択されたモデル名を取得
                selected_model_name = self.selected_gemini_model.get()
                logger.info(f"使用する Gemini モデル: {selected_model_name}")
                if self.gemini_client is None:
                    raise Exception("Gemini API クライアントが設定されていません。")
                response = self.gemini_client.models.generate_content(
                    model=selected_model_name,  # 選択されたモデルを使用
                    contents=[video_file, prompt],
                    config={
                        "response_mime_type": "application/json",
                        "response_schema": schema,
                    },
                )
                return response
            except Exception as e:
                logger.warning(f"APIサーバーエラー（試行 {attempt+1}）: {e}")
                if attempt < max_retries - 1:
                    logger.info(f"{retry_delay}秒後に再試行します...")
                    time.sleep(retry_delay)
                else:
                    logger.error(f"最大試行回数に達しました。エラー: {e}")
                    raise

    def finish_note_creation(self, success, result_path_or_error):
        """ノート作成処理の完了をUIに反映する"""
        if success:
            # ★ 成功時のステータス
            self.note_creation_status.set("ノート生成完了")
            self.note_result_message.set(
                f"保存先: {os.path.basename(result_path_or_error)}"
            )  # ファイル名のみ表示
            logger.info(f"ノート作成成功: {result_path_or_error}")
            messagebox.showinfo(
                "ノート作成完了",
                f"ノートが正常に作成されました。\n{result_path_or_error}",
            )
        else:
            # ★ 失敗時のステータス
            self.note_creation_status.set("ノート生成失敗")
            # エラーメッセージを短縮して表示
            error_short = str(result_path_or_error).split("\n")[0][:100] + "..."
            self.note_result_message.set(
                f"エラー: {error_short}"
            )  # 修正: error_short を使用
            logger.error(f"ノート作成失敗: {result_path_or_error}")
            messagebox.showerror(
                "ノート作成エラー",
                f"ノートの作成中にエラーが発生しました:\n{result_path_or_error}",
            )

        # --- GUI操作制限解除 ---
        self._restore_gui_after_note_creation()

    def stop_all_tasks(self):
        """全てのキャプチャ・録画タスクを停止する"""
        logger.info("全てのタスクの停止処理を開始します。")

        # ★★★ 修正点: 停止ボタンをすぐに無効化 ★★★
        self.stop_button.config(state=tk.DISABLED)
        # 開始ボタンは、関連処理がすべて完了するまで無効のままにする
        # (finish_note_creation またはこの関数の最後で有効に戻す)

        was_recording = self.is_recording  # 録画中だったか記録

        # 録画停止 (音声も内部で停止される)
        if self.is_recording:
            self.stop_recording()  # stop_recording が呼ばれる (内部で保存とノート作成トリガー)

        # スクリーンショット停止
        if self.is_capturing_screenshot:
            self.stop_screenshot_capture()

        if was_recording:
            logger.info(
                "録画を停止しました。動画保存とノート作成がバックグラウンドで実行されます（成功した場合）。"
            )
            # この場合、UIの有効化は finish_note_creation に任せる
            # finish_note_creation が呼ばれないケース(動画保存失敗など)も考慮すると、
            # ここで start_button などを有効にすべきではない。
        else:
            logger.info("スクリーンショットキャプチャを停止しました。")
            # スクリーンショットのみ停止した場合、ノート作成は走らないので、
            # ここでUIを操作可能に戻す。
            self._update_start_button_state()
            self.folder_entry.config(state=tk.NORMAL)
            self.refresh_window_list_button.config(state=tk.NORMAL)
            self.window_listbox.config(state=tk.NORMAL)
            self.model_combobox.config(state="readonly")
            self._set_checkbox_state(tk.NORMAL)
            # スリープ設定を元に戻す
            try:
                self._restore_sleep()
            except Exception:
                pass
            # 閉じるボタンの挙動も元に戻す (ノート作成がない場合)
            if hasattr(self, "original_on_closing"):
                self.root.protocol("WM_DELETE_WINDOW", self.original_on_closing)
            else:
                self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
            # スクリーンショットのみ停止時も設定タブを解放
            self._unlock_settings_tab()

        logger.info("全てのタスクの停止処理を完了しました。")

    def prepare_save_folder(self):
        """保存フォルダの準備（作成、権限チェック）を行う"""
        initial_folder_name_from_ui = self.save_folder_name.get()
        current_working_directory = os.getcwd()  # ★ 現在の作業ディレクトリをログに出力
        logger.info(
            f"prepare_save_folder - UIから取得した初期フォルダ名: '{initial_folder_name_from_ui}'"
        )
        logger.info(
            f"prepare_save_folder - 現在の作業ディレクトリ: {current_working_directory}"
        )

        folder_name = initial_folder_name_from_ui
        if not folder_name:
            # フォルダ名が空の場合、現在時刻でデフォルト名を生成
            folder_name = f"capture_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            # self.save_folder_name.set(folder_name) # UIへのセットは成功後の方が良いかもしれない
            logger.info(
                f"prepare_save_folder - 保存フォルダ名が指定されなかったため、デフォルト名を生成しました: {folder_name}"
            )

        # folder_name が絶対パスか相対パスかで処理を分ける
        if os.path.isabs(folder_name):
            save_path = folder_name
            logger.info(
                f"prepare_save_folder - 入力されたフォルダ名は絶対パスとして扱います: {save_path}"
            )
        else:
            # 相対パスの場合、現在の作業ディレクトリを基準に絶対パスを生成
            save_path = os.path.join(current_working_directory, folder_name)
            # os.path.abspath は上記とほぼ同等だが、より明示的に join を使う
            # save_path = os.path.abspath(folder_name)
            logger.info(
                f"prepare_save_folder - 相対パスとして解釈し、絶対パスに変換しました (基準: {current_working_directory}): {save_path}"
            )

        logger.info(
            f"prepare_save_folder - 最終的な保存先フォルダ (os.makedirs対象): {save_path}"
        )

        try:
            # フォルダが存在しない場合は作成
            if not os.path.exists(save_path):
                logger.info(
                    f"prepare_save_folder - os.makedirs を呼び出します: Path='{save_path}', exist_ok=True"
                )
                os.makedirs(save_path, exist_ok=True)
                # makedirs がエラーを出さなかった場合、通常は作成されているはず
                logger.info(
                    f"prepare_save_folder - フォルダを作成しました (または既に存在していました): {save_path}"
                )
                if os.path.exists(save_path):
                    logger.info(
                        f"prepare_save_folder - フォルダ作成後の存在確認OK: {save_path}"
                    )
                else:
                    # このケースは稀だが、makedirsが成功したように見えても実際には作成されていない場合
                    logger.error(
                        f"prepare_save_folder - フォルダ作成後に存在確認NG (os.makedirsはエラーなし): {save_path}"
                    )
            else:
                logger.info(
                    f"prepare_save_folder - フォルダは既に存在します: {save_path}"
                )

            # 書き込み権限をチェック (簡易的な方法)
            test_file_path = os.path.join(save_path, ".permission_test")
            logger.info(
                f"prepare_save_folder - 書き込み権限テストファイルを作成します: {test_file_path}"
            )
            with open(test_file_path, "w") as f:
                f.write("test")
            os.remove(test_file_path)
            logger.info(
                f"prepare_save_folder - フォルダへの書き込み権限を確認しました: {save_path}"
            )
            self.save_folder_name.set(
                save_path
            )  # フォルダ準備成功後にUIに絶対パスを反映
            return True

        except OSError as e:
            # ★ OSError の詳細情報をログに出力
            logger.exception(
                f"prepare_save_folder - 保存フォルダの準備中にOSErrorが発生しました: Path='{save_path}', ErrorNo={e.errno}, ErrorMsg='{e.strerror}', Details='{e}'"
            )
            messagebox.showerror(
                "フォルダエラー",
                f"保存フォルダの作成またはアクセスに失敗しました:\n{save_path}\n\nエラーコード: {e.errno}\n詳細: {e.strerror}",
            )
            return False
        except Exception as e:
            logger.exception(
                f"prepare_save_folder - 保存フォルダの準備中に予期せぬエラーが発生しました: Path='{save_path}', Error='{e}'"
            )
            messagebox.showerror(
                "フォルダエラー",
                f"保存フォルダの準備中に予期せぬエラーが発生しました:\n{save_path}\n\nエラー詳細:\n{e}",
            )
            return False

    # --- ステータス更新メソッド ---
    def update_screenshot_status(self):
        """スクリーンショットのステータス表示を更新する"""
        if self.is_capturing_screenshot:
            status_text = f"実行中... ({self.screenshot_saved_count}枚保存)"
            if self.last_saved_screenshot_filename:
                status_text += f"\n最終保存: {os.path.basename(self.last_saved_screenshot_filename)}"
            self.screenshot_status_label.config(text=status_text)
            # 1秒後に再度更新
            self.root.after(1000, self.update_screenshot_status)
        else:
            status_text = "停止中"
            if self.last_saved_screenshot_filename:
                status_text += f" (最終保存: {os.path.basename(self.last_saved_screenshot_filename)})"
            elif self.screenshot_saved_count > 0:
                status_text += f" ({self.screenshot_saved_count}枚保存)"
            else:
                status_text = "待機中..."
            self.screenshot_status_label.config(text=status_text)

    def update_recording_status(self):
        """録画のステータス表示を更新する"""
        if self.is_recording and self.recording_start_time:
            elapsed_seconds = int(time.time() - self.recording_start_time)
            minutes, seconds = divmod(elapsed_seconds, 60)
            status_text = f"録画中... {minutes:02d}:{seconds:02d}"
            if self.last_saved_video_filename:  # 保存が完了していれば表示
                status_text += (
                    f"\n最終保存: {os.path.basename(self.last_saved_video_filename)}"
                )
            self.recording_status_label.config(text=status_text)
            # 1秒後に再度更新
            self.root.after(1000, self.update_recording_status)
        else:
            status_text = "停止中"
            if self.last_saved_video_filename:
                status_text += (
                    f" (最終保存: {os.path.basename(self.last_saved_video_filename)})"
                )
            else:
                status_text = "待機中..."
            self.recording_status_label.config(text=status_text)

    # --- スクリーンショット関連メソッド ---
    def start_screenshot_capture(self):
        """スクリーンショットの連続キャプチャを開始する"""
        if self.is_capturing_screenshot:
            return

        hwnd = self.selected_window_handle.get()  # 選択されたウィンドウハンドルを取得

        logger.info(
            f"スクリーンショットキャプチャを開始します。対象ウィンドウハンドル: {hwnd if hwnd != 0 else 'プライマリモニター'}"
        )
        # フォルダ準備は start_tasks で行う

        # スクリーンショットキャプチャスレッドを開始 (hwnd を渡す)
        self.screenshot_thread = threading.Thread(
            target=self.screenshot_capture_loop,
            args=(hwnd,),
            name="ScreenshotThread",
            daemon=True,
        )
        self.is_capturing_screenshot = True
        self.screenshot_saved_count = 0  # カウンタリセット
        self.last_saved_screenshot_filename = ""  # ファイル名リセット
        self.screenshot_thread.start()

        # self.start_screenshot_button.config(state=tk.DISABLED) # 個別ボタン削除
        # self.stop_screenshot_button.config(state=tk.NORMAL) # 個別ボタン削除
        # ボタン状態は start_tasks / stop_all_tasks で制御
        self.update_screenshot_status()  # ステータス更新開始

    def stop_screenshot_capture(self):
        """スクリーンショットの連続キャプチャを停止する"""
        if not self.is_capturing_screenshot:
            return

        logger.info("スクリーンショットキャプチャの停止処理を開始します。")
        self.is_capturing_screenshot = False
        if self.screenshot_thread and self.screenshot_thread.is_alive():
            logger.info("スクリーンショットキャプチャスレッドの終了を待機します...")
            self.screenshot_thread.join(timeout=5.0)  # タイムアウトを設定
            if self.screenshot_thread.is_alive():
                logger.warning(
                    "スクリーンショットキャプチャスレッドが時間内に終了しませんでした。"
                )
            else:
                logger.info("スクリーンショットキャプチャスレッドが終了しました。")
        self.screenshot_thread = None

        # self.start_screenshot_button.config(state=tk.NORMAL) # 個別ボタン削除
        # self.stop_screenshot_button.config(state=tk.DISABLED) # 個別ボタン削除
        # ボタン状態は start_tasks / stop_all_tasks で制御
        self.update_screenshot_status()  # ステータス更新

        # エラーチェック
        if self.error_occurred_in_screenshot_thread:
            messagebox.showerror(
                "スクリーンショットエラー",
                "スクリーンショット取得中にエラーが発生しました。詳細はログを確認してください。",
            )
            self.error_occurred_in_screenshot_thread = False  # フラグをリセット

        logger.info("スクリーンショットキャプチャの停止処理を完了しました。")

    def screenshot_capture_loop(self, hwnd):  # 引数 hwnd を追加
        """指定されたウィンドウまたは画面全体のスクリーンショットを定期的に取得し、変化があれば保存するループ"""
        logger.info(
            f"スクリーンショットキャプチャループを開始します。対象ウィンドウハンドル: {hwnd if hwnd != 0 else 'プライマリモニター'}"
        )
        interval_seconds = 1  # 取得間隔（秒）
        self.last_screenshot_image = None  # 前回の画像をリセット
        self.last_saved_screenshot_image = None  # 最後に保存した画像をリセット

        try:  # sct オブジェクトの初期化を try の外に出す
            sct = mss.mss()
        except Exception as e:
            logger.exception(f"mss の初期化中にエラー: {e}")
            self.error_occurred_in_screenshot_thread = True
            return  # 初期化失敗時はループに入らない

        while self.is_capturing_screenshot:
            start_capture_time = time.time()
            try:
                # キャプチャ領域を決定
                monitor = None
                current_hwnd = hwnd  # ループ内で hwnd を変更する可能性があるのでコピー
                if current_hwnd != 0:
                    try:
                        # ウィンドウが存在するか確認
                        if not win32gui.IsWindow(current_hwnd):
                            logger.warning(
                                f"ウィンドウハンドル {current_hwnd} が無効です。プライマリモニターを対象にします。"
                            )
                            current_hwnd = (
                                0  # 無効ならプライマリモニターにフォールバック
                            )
                        else:
                            window_rect = win32gui.GetWindowRect(current_hwnd)
                            left, top, right, bottom = window_rect
                            width = right - left
                            height = bottom - top
                            if width > 0 and height > 0:
                                monitor = {
                                    "left": left,
                                    "top": top,
                                    "width": width,
                                    "height": height,
                                }
                                # logger.debug(f"対象ウィンドウ領域: {monitor}") # デバッグ用
                            else:
                                logger.warning(
                                    f"ウィンドウサイズが無効です ({width}x{height})。プライマリモニターを対象にします。"
                                )
                                current_hwnd = 0  # ウィンドウが無効ならプライマリモニターにフォールバック
                    except Exception as e:
                        logger.warning(
                            f"ウィンドウ情報の取得に失敗しました: {e}。プライマリモニターを対象にします。"
                        )
                        current_hwnd = 0  # エラー時もプライマリモニターにフォールバック

                if (
                    monitor is None
                ):  # current_hwnd が 0 またはウィンドウ情報取得失敗の場合
                    # プライマリモニターを取得 (monitors[0] は全画面結合の場合があるので注意、通常は monitors[1])
                    # 以前の実装 sct.monitors[0] に合わせるか、録画処理 sct.monitors[1] に合わせるか
                    # ここでは monitors[1] を試す (より一般的)
                    if len(sct.monitors) > 1:
                        monitor = sct.monitors[1]
                    else:
                        monitor = sct.monitors[0]  # モニターが1つしかない場合
                    # logger.debug(f"対象領域: プライマリモニター {monitor}") # デバッグ用

                # スクリーンショットを取得
                screenshot = sct.grab(monitor)
                # Pillowイメージに変換
                current_image_pil = Image.frombytes(
                    "RGB", screenshot.size, screenshot.rgb
                )
                # OpenCV形式に変換 (比較用)
                current_image_cv = cv2.cvtColor(
                    np.array(current_image_pil), cv2.COLOR_RGB2BGR
                )

                if self.last_saved_screenshot_image is not None:
                    # 最後に保存した画像と比較して変化があるか確認
                    if not self.is_similar(
                        current_image_cv, self.last_saved_screenshot_image
                    ):
                        logger.info(
                            "画面に変化を検出しました。スクリーンショットを保存します。"
                        )
                        if self.save_screenshot_image(
                            current_image_cv
                        ):  # 保存処理を呼び出し
                            self.screenshot_saved_count += 1
                    # else:
                    #     logger.debug("画面に変化はありません。") # デバッグ用
                else:
                    # 最初のスクリーンショットは必ず保存
                    logger.info("最初のスクリーンショットを保存します。")
                    if self.save_screenshot_image(current_image_cv):
                        self.screenshot_saved_count += 1

                # 今回の画像を保持（デバッグ用など、必要に応じて）
                self.last_screenshot_image = current_image_cv

            except UnidentifiedImageError as e:
                logger.error(f"スクリーンショット画像の形式が認識できませんでした: {e}")
                self.error_occurred_in_screenshot_thread = True
                # エラーが発生してもループは継続するかもしれないが、一旦フラグを立てる
            except mss.ScreenShotError as sct_err:  # mss のエラーもキャッチ
                logger.error(
                    f"スクリーンショット取得中にエラーが発生しました: {sct_err}"
                )
                self.error_occurred_in_screenshot_thread = True
            except Exception as e:
                logger.exception(
                    f"スクリーンショット取得/比較中にエラーが発生しました: {e}"
                )
                self.error_occurred_in_screenshot_thread = True
                # エラーが発生してもループは継続するかもしれないが、一旦フラグを立てる

            # 次のキャプチャまでの待機時間
            elapsed_time = time.time() - start_capture_time
            sleep_time = max(0, interval_seconds - elapsed_time)
            if sleep_time > 0:
                time.sleep(sleep_time)

        logger.info("スクリーンショットキャプチャループを終了します。")

    def is_similar(self, img1_cv, img2_cv, threshold=None):
        """2つの画像の類似度を計算する (差分ベースの簡易比較)"""
        try:
            # 設定から閾値を取得（引数で指定されていない場合）
            if threshold is None:
                threshold = self.similarity_threshold

            # グレースケールに変換

            # サイズが異なる場合はリサイズ (小さい方に合わせるか、固定サイズにする)
            if img1_cv.shape != img2_cv.shape:
                # 例: img1 のサイズに合わせる
                h, w = img1_cv.shape[:2]
                img2_cv = cv2.resize(img2_cv, (w, h))
                logger.warning("比較画像のサイズが異なるためリサイズしました。")

            # 差分を計算
            diff = cv2.absdiff(img1_cv, img2_cv)
            diff_mean = np.mean(diff, axis=2)  # 各ピクセルのRGB差分の平均を計算

            # 差分が閾値以下のピクセルの割合を計算（設定から取得した閾値を使用）
            non_zero_count = np.count_nonzero(diff_mean > self.diff_pixel_threshold)
            total_pixels = diff.shape[0] * diff.shape[1]
            similarity = 1.0 - (non_zero_count / total_pixels)

            logger.debug(f"画像類似度: {similarity:.4f}")  # デバッグ用

            return similarity >= threshold
        except cv2.error as e:
            logger.error(f"画像比較 (is_similar) 中にOpenCVエラーが発生しました: {e}")
            return False  # エラー時は類似していないと判断
        except Exception as e:
            logger.exception(
                f"画像比較 (is_similar) 中に予期せぬエラーが発生しました: {e}"
            )
            return False  # エラー時は類似していないと判断

    def save_screenshot_image(self, image_cv):  # save_image から変更
        """スクリーンショット画像をファイルに保存する。エラー発生時はログに記録。

        日本語パス対応: cv2.imwrite の代わりに cv2.imencode を使用して
        Pythonのファイル操作で保存することで、日本語を含むパスに対応しています。
        """
        save_folder = self.save_folder_name.get()
        if not save_folder:
            logger.error(
                "保存フォルダが設定されていません。スクリーンショットを保存できません。"
            )
            return False

        filepath = None
        try:
            now = datetime.now()
            # ミリ秒を含むファイル名
            filename = f"screenshot_{now.strftime('%Y%m%d_%H%M%S')}_{now.microsecond // 1000:03d}.png"
            filepath = os.path.join(save_folder, filename)

            # 日本語パス対応: cv2.imencode でメモリ上にエンコードし、
            # Pythonのファイル操作で保存（日本語パスに対応）
            success, buffer = cv2.imencode(".png", image_cv)

            if success:
                # バイナリモードでファイルに書き込み
                with open(filepath, "wb") as f:
                    f.write(buffer)

                logger.info(f"スクリーンショットを保存しました: {filepath}")
                self.last_saved_screenshot_filename = (
                    filepath  # 最後に保存したファイル名を更新
                )
                # 最後に保存した画像を比較用に保持（コピーして保存）
                self.last_saved_screenshot_image = image_cv.copy()
                return True
            else:
                logger.error(
                    f"スクリーンショットのエンコードに失敗しました (cv2.imencodeがFalseを返しました): {filepath}"
                )
                return False
        except cv2.error as e:
            logger.error(
                f"スクリーンショット保存中にOpenCVエラーが発生しました: {e}. ファイルパス: {filepath if 'filepath' in locals() else 'N/A'}"
            )
            return False
        except IOError as e:
            logger.error(
                f"スクリーンショット保存中にファイルI/Oエラーが発生しました: {e}. ファイルパス: {filepath if 'filepath' in locals() else 'N/A'}"
            )
            return False
        except Exception as e:
            logger.exception(
                f"スクリーンショット保存中に予期せぬエラーが発生しました: {e}. ファイルパス: {filepath if 'filepath' in locals() else 'N/A'}"
            )
            return False

    def on_closing(self):
        """ウィンドウが閉じられるときの処理"""
        logger.info("アプリケーション終了処理を開始します。")
        if self.is_capturing_screenshot or self.is_recording:
            if messagebox.askokcancel(
                "確認", "キャプチャまたは録画が実行中です。本当に終了しますか？"
            ):
                self.stop_all_tasks()  # 実行中のタスクを停止
                # stop_all_tasks内で録画停止→保存→ノート作成が走る可能性があるため、少し待つか、
                # ノート作成完了まで待つ仕組みが必要かもしれない。
                # 現状では、ノート作成が完了する前にウィンドウが閉じる可能性がある。
                logger.info("実行中のタスクを停止しました。")
                self.root.destroy()
            else:
                logger.info("アプリケーションの終了をキャンセルしました。")
                return  # 終了をキャンセル
        else:
            self.root.destroy()
        logger.info("アプリケーションを終了しました。")


if __name__ == "__main__":
    try:
        try:

            # Per-Monitor DPI Aware V2 (Windows 10 Creators Update以降)
            # windll.shcore.SetProcessDpiAwareness(2)
            # System DPI Aware (より古いWindowsや互換性重視の場合)
            # windll.shcore.SetProcessDpiAwareness(1) # 既存のコードではこちらが使われている可能性

            # ★★★ 変更箇所 スタート ★★★
            # Per-Monitor DPI Aware V2 に設定してみる
            # これにより、各モニターのDPI設定が個別に認識されるようになる
            PROCESS_PER_MONITOR_DPI_AWARE = 2
            windll.shcore.SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE)
            logger.info("DPI Awareness を Per-Monitor V2 (2) に設定しました。")
            # ★★★ 変更箇所 エンド ★★★

        except ImportError:
            logger.info(
                "ctypesモジュールが見つからないため、DPI Awareness は設定されませんでした (Windows以外の環境の可能性)。"
            )
        except AttributeError:
            # Per-Monitor V2 が利用できない古いWindowsの場合、System Aware を試す
            try:
                PROCESS_SYSTEM_DPI_AWARE = 1
                windll.shcore.SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
                logger.info(
                    "DPI Awareness を System Aware (1) に設定しました (Per-Monitor V2 不可のため)。"
                )
            except AttributeError:
                logger.warning(
                    "DPI Awareness の設定に失敗しました (古いWindowsバージョンの可能性)。"
                )
            except Exception as e_sys:
                logger.warning(f"System DPI Awareness 設定中にエラー: {e_sys}")
        except Exception as e:
            logger.warning(f"DPI Awareness 設定中に予期せぬエラー: {e}")

        root = tk.Tk()
        photo = tk.PhotoImage(file=resource_path("icon", "app_icon.png"))
        root.iconphoto(True, photo)
        app = SlideCaptureApp(root)
        root.mainloop()
    except Exception as e:
        # GUI初期化前などのエラーをキャッチ
        logger.critical(
            f"アプリケーションの起動中に致命的なエラーが発生しました: {e}",
            exc_info=True,
        )
        # コンソールにもエラーを表示
        print(f"致命的なエラーが発生しました: {e}")
        traceback.print_exc()
        # 簡単なメッセージボックスでユーザーに通知 (Tkinterが使える場合)
        try:
            error_root = tk.Tk()
            error_root.withdraw()  # メインウィンドウは表示しない
            messagebox.showerror(
                "起動エラー",
                f"アプリケーションの起動に失敗しました。\n詳細はログファイルを確認してください。\n\n{e}",
            )
            error_root.destroy()
        except tk.TclError:
            print(
                "メッセージボックスを表示できませんでした。"
            )  # Tkinterが初期化失敗した場合
        except Exception as msg_e:
            print(f"エラーメッセージ表示中にさらにエラー: {msg_e}")
