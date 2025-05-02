# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
from datetime import datetime
import os
import logging  # logging モジュールをインポート
import traceback  # スタックトレース取得のため

import win32gui  # ウィンドウ操作のため追加

import mss  # 画面キャプチャのため追加
import cv2
from PIL import Image, ImageGrab, UnidentifiedImageError
import numpy as np

# --- ロギング設定 ---
log_filename = "slide_capture_app.log"
log_format = "%(asctime)s - %(levelname)s - %(threadName)s - %(message)s"
logging.basicConfig(
    level=logging.INFO,
    format=log_format,
    handlers=[
        logging.FileHandler(log_filename, encoding="utf-8"),  # ファイル出力
        logging.StreamHandler(),  # コンソールにも出力 (デバッグ用)
    ],
)
logger = logging.getLogger(__name__)


class SlideCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("スライドキャプチャ＆録画")
        # UIの高さを増やして新しい要素を配置
        self.root.geometry("450x550")  # サイズ調整

        # --- 状態変数 ---
        self.is_capturing_screenshot = False  # スクリーンショット中フラグ
        self.is_recording = False  # 録画中フラグ
        self.screenshot_thread = None
        self.recording_thread = None
        self.recorded_frames = []  # 録画フレームを一時保存するリスト
        self.last_screenshot_image = None
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

        # --- UI要素の作成 ---

        # --- 保存フォルダ設定 ---
        folder_frame = ttk.Frame(root, padding="10")
        folder_frame.pack(fill=tk.X)
        folder_label = ttk.Label(folder_frame, text="保存フォルダ名:")
        folder_label.pack(side=tk.LEFT, padx=(0, 5))
        self.folder_entry = ttk.Entry(
            folder_frame, textvariable=self.save_folder_name, width=45  # 幅調整
        )
        self.folder_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # --- ウィンドウ選択 ---
        window_selection_frame = ttk.LabelFrame(
            root, text="録画対象ウィンドウ選択", padding="10"
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

        # --- 操作ボタン (統合) ---
        button_frame = ttk.Frame(root, padding="10")
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

        # --- ステータス表示 ---
        status_frame = ttk.Frame(root, padding="10")
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

        logger.info("アプリケーションを初期化しました。")
        self.refresh_window_list()  # 初期ウィンドウリスト表示

    # --- ウィンドウ選択関連メソッド ---
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

    # --- 統合開始・停止メソッド ---
    def start_tasks(self):
        """スクリーンショットと録画を開始する（ウィンドウ選択状態による）"""
        if self.is_capturing_screenshot or self.is_recording:
            logger.warning("既にタスクが実行中です。")
            return

        hwnd = self.selected_window_handle.get()

        # フォルダ設定は共通で最初に行う
        if not self.prepare_save_folder():
            return

        # スクリーンショットは常に開始
        self.start_screenshot_capture()

        # ウィンドウが選択されていれば録画も開始
        if hwnd != 0:
            self.start_recording(hwnd)  # 引数でハンドルを渡す
        else:
            logger.info(
                "ウィンドウが選択されていないため、スクリーンショットのみ開始します。"
            )

        # 統合ボタンの状態更新
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.folder_entry.config(state=tk.DISABLED)
        self.refresh_window_list_button.config(
            state=tk.DISABLED
        )  # タスク実行中は更新不可
        self.window_listbox.config(state=tk.DISABLED)

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
        # ここで選択されたウィンドウハンドル(hwnd)を使い、録画スレッドを開始する（未実装）
        self.is_recording = True
        # self.start_recording_button.config(state=tk.DISABLED) # 個別ボタン削除
        # self.stop_recording_button.config(state=tk.NORMAL) # 個別ボタン削除
        # ボタン状態は start_tasks / stop_all_tasks で制御
        self.recording_start_time = time.time()
        self.update_recording_status()  # ステータス更新開始

    def stop_recording(self):
        """録画処理を停止する"""
        if not self.is_recording:
            return

        logger.info("録画停止処理を開始します。（未実装）")
        self.is_recording = False
        # ここで録画スレッドを停止し、動画ファイルを処理する（未実装）
        # self.start_recording_button.config(state=tk.NORMAL) # 個別ボタン削除
        # self.stop_recording_button.config(state=tk.DISABLED) # 個別ボタン削除
        # ボタン状態は stop_all_tasks で制御
        self.update_recording_status()  # 最終ステータス表示
        # 録画停止時にノート作成を開始
        self.start_note_creation()  # ノート作成は録画停止時にトリガー

    def start_note_creation(self):
        logger.info("ノート作成処理（未実装）")
        self.note_creation_status.set("ノート作成中...")
        self.note_result_message.set("")
        # ここで Gemini API 連携と Word ファイル生成を行うスレッドを開始する
        # 仮の完了処理
        self.root.after(
            3000, self.finish_note_creation, True, "ノート作成完了: result.docx"
        )

    def finish_note_creation(self, success, message):
        logger.info(f"ノート作成完了: success={success}, message={message}")
        if success:
            self.note_creation_status.set("ノート作成完了")
            self.note_result_message.set(f"保存先: {message}")
        else:
            self.note_creation_status.set("ノート作成エラー")
            self.note_result_message.set(f"エラー: {message}")

    def stop_all_tasks(self):
        """実行中の全てのタスク（スクリーンショット、録画）を停止する"""
        if not self.is_capturing_screenshot and not self.is_recording:
            return

        logger.info("すべてのタスクの停止を試みます...")
        self.stop_screenshot_capture()
        self.stop_recording()  # stop_recording はノート作成をトリガーする可能性があるので注意

        # 統合ボタンの状態更新
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.folder_entry.config(state=tk.NORMAL)
        self.refresh_window_list_button.config(state=tk.NORMAL)
        self.window_listbox.config(state=tk.NORMAL)
        logger.info("すべてのタスクを停止しました。")

    # --- フォルダ準備メソッド (start_captureから分離) ---
    def prepare_save_folder(self):
        """保存フォルダの準備（作成、権限チェック）を行う"""
        folder_name_input = self.save_folder_name.get().strip()
        if not folder_name_input:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            folder_name = f"capture_{timestamp}"  # デフォルト名を変更
            self.save_folder_name.set(folder_name)
            logger.info(f"フォルダ名が未入力のため、デフォルト名を設定: {folder_name}")
        else:
            # ファイル名として不適切な文字を置換 (簡易的な対策)
            invalid_chars = '<>:"/\\|?*'
            folder_name = "".join(
                c if c not in invalid_chars else "_" for c in folder_name_input
            )
            if folder_name != folder_name_input:
                self.save_folder_name.set(folder_name)
                warning_msg = f"フォルダ名に使用できない文字が含まれていたため、'{folder_name}' に修正しました。"
                logger.warning(warning_msg)
                messagebox.showwarning("フォルダ名修正", warning_msg)

        try:
            self.capture_save_path = os.path.abspath(folder_name)
            logger.info(f"保存先フォルダの絶対パス: {self.capture_save_path}")

            if not os.path.exists(self.capture_save_path):
                logger.info(
                    f"フォルダが存在しないため作成します: {self.capture_save_path}"
                )
                os.makedirs(self.capture_save_path, exist_ok=True)
                logger.info(f"フォルダを作成しました: {self.capture_save_path}")
            elif not os.path.isdir(self.capture_save_path):
                error_msg = (
                    f"指定されたパスはフォルダではありません: {self.capture_save_path}"
                )
                logger.error(error_msg)
                messagebox.showerror("エラー", error_msg)
                return False
            # 書き込み権限チェック (Windows用簡易チェック)
            elif not os.access(self.capture_save_path, os.W_OK):
                error_msg = f"指定されたフォルダへの書き込み権限がありません:\n{self.capture_save_path}\n別のフォルダを指定するか、権限を確認してください。"
                logger.error(error_msg)
                messagebox.showerror("権限エラー", error_msg)
                return False
            else:
                logger.info(
                    f"保存先フォルダへの書き込み権限を確認しました: {self.capture_save_path}"
                )
            return True  # 成功

        except OSError as e:
            error_detail = (
                f"フォルダの作成/アクセス中にOSエラーが発生しました。\n"
                f"エラータイプ: {type(e).__name__}\n"
                f"エラーコード: {e.errno}\n"
                f"メッセージ: {e.strerror}\n"
                f"パス: {getattr(e, 'filename', 'N/A')}"
            )
            logger.error(f"{error_detail}\n{traceback.format_exc()}")
            messagebox.showerror(
                "フォルダエラー",
                f"{error_detail}\n詳細はログファイル ({log_filename}) を確認してください。",
            )
            return False
        except Exception as e:
            error_detail = (
                f"フォルダ処理中に予期せぬエラーが発生しました:\n"
                f"エラータイプ: {type(e).__name__}\n"
                f"メッセージ: {e}"
            )
            logger.exception(error_detail)
            messagebox.showerror(
                "予期せぬエラー",
                f"{error_detail}\n詳細はログファイル ({log_filename}) を確認してください。",
            )
            return False

    # --- ステータス更新メソッド ---
    def update_screenshot_status(self):
        """スクリーンショットのステータスラベルを更新する"""
        if self.is_capturing_screenshot:
            status_text = f"実行中...\n保存枚数: {self.screenshot_saved_count}\n最終保存: {self.last_saved_screenshot_filename}"
            if self.error_occurred_in_screenshot_thread:
                status_text += f"\n警告: エラー発生。詳細はログファイル\n({log_filename})を確認してください。"
                self.screenshot_status_label.config(foreground="red")
            else:
                self.screenshot_status_label.config(foreground="black")
            self.screenshot_status_label.config(text=status_text)
            if self.is_capturing_screenshot:
                self.root.after(1000, self.update_screenshot_status)
        else:
            final_status = f"停止中。\n合計保存枚数: {self.screenshot_saved_count}"
            if self.last_saved_screenshot_filename:
                final_status += f"\n最終保存: {self.last_saved_screenshot_filename}"
            if self.error_occurred_in_screenshot_thread:
                final_status += f"\n警告: エラー発生。詳細はログファイル\n({log_filename})を確認してください。"
                self.screenshot_status_label.config(foreground="red")
            else:
                self.screenshot_status_label.config(foreground="black")
            self.screenshot_status_label.config(text=final_status)

    def update_recording_status(self):
        """録画のステータスラベルを更新する"""
        if self.is_recording:
            elapsed_time = time.time() - self.recording_start_time
            status_text = (
                f"録画中... ({int(elapsed_time // 60)}:{int(elapsed_time % 60):02d})"
            )
            if self.error_occurred_in_recording_thread:
                status_text += f"\n警告: エラー発生。詳細はログファイル\n({log_filename})を確認してください。"
                self.recording_status_label.config(foreground="red")
            else:
                self.recording_status_label.config(foreground="black")
            self.recording_status_label.config(text=status_text)
            if self.is_recording:
                self.root.after(1000, self.update_recording_status)
        else:
            final_status = "停止中。"
            if self.last_saved_video_filename:
                final_status += f"\n最終保存: {self.last_saved_video_filename}"
            if self.error_occurred_in_recording_thread:
                final_status += f"\n警告: エラー発生。詳細はログファイル\n({log_filename})を確認してください。"
                self.recording_status_label.config(foreground="red")
            else:
                self.recording_status_label.config(foreground="black")
            self.recording_status_label.config(text=final_status)

    def start_screenshot_capture(self):
        """スクリーンショットキャプチャ処理を開始する"""
        # フォルダ準備は start_tasks で行うため削除
        if self.is_capturing_screenshot:
            return

        # エラーフラグをリセット
        self.error_occurred_in_screenshot_thread = (
            False  # error_occurred_in_thread から変更
        )
        self.screenshot_status_label.config(
            foreground="black"
        )  # ステータス色をリセット

        self.is_capturing_screenshot = True
        # self.start_screenshot_button.config(state=tk.DISABLED) # 個別ボタン削除
        # self.stop_screenshot_button.config(state=tk.NORMAL) # 個別ボタン削除
        # ボタン状態は start_tasks / stop_all_tasks で制御
        # self.folder_entry.config(state=tk.DISABLED) # start_tasks で制御
        self.screenshot_saved_count = 0
        self.last_saved_screenshot_filename = ""
        self.last_screenshot_image = None  # last_image から変更

        # スレッドを開始
        self.screenshot_thread = threading.Thread(  # capture_thread から変更
            target=self.screenshot_capture_loop,
            name="ScreenshotCaptureThread",
            daemon=True,  # capture_loop から変更
        )
        self.screenshot_thread.start()  # capture_thread から変更
        self.update_screenshot_status()  # 定期的なステータス更新を開始 (update_status から変更)
        logger.info(
            f"スクリーンショットキャプチャを開始しました。保存先: {self.capture_save_path}"
        )

    def stop_screenshot_capture(self):
        """スクリーンショットキャプチャ処理を停止する"""
        if not self.is_capturing_screenshot:
            return

        logger.info("スクリーンショットキャプチャ停止処理を開始します。")
        self.is_capturing_screenshot = False  # is_capturing から変更
        # スレッドが終了するのを少し待つ
        if (
            self.screenshot_thread and self.screenshot_thread.is_alive()
        ):  # capture_thread から変更
            logger.info("スクリーンショットキャプチャループの終了を待っています...")
            self.screenshot_thread.join(
                timeout=1.5
            )  # 少し長めに待つ # capture_thread から変更
            if self.screenshot_thread.is_alive():  # capture_thread から変更
                logger.warning(
                    "警告: スクリーンショットキャプチャスレッドが時間内に終了しませんでした。"
                )

        # self.start_screenshot_button.config(state=tk.NORMAL) # 個別ボタン削除
        # self.stop_screenshot_button.config(state=tk.DISABLED) # 個別ボタン削除
        # self.folder_entry.config(state=tk.NORMAL) # stop_all_tasks で制御
        # logger.info("スクリーンショットキャプチャを停止しました。") # stop_all_tasks でログ出力
        # 停止後にもう一度ステータスを更新して最終結果を表示
        self.update_screenshot_status()  # 即時更新に変更

    def screenshot_capture_loop(self):
        """定期的にスクリーンショットを取得し、比較・保存するループ"""
        logger.info("スクリーンショットキャプチャループを開始します。")
        while self.is_capturing_screenshot:  # is_capturing から変更
            try:
                # 1. スクリーンショット取得
                screenshot = ImageGrab.grab()
                if screenshot is None:
                    logger.error(
                        "エラー: ImageGrab.grab() が None を返しました。スクリーンショットを取得できませんでした。"
                    )
                    self.error_occurred_in_screenshot_thread = (
                        True  # error_occurred_in_thread から変更
                    )
                    time.sleep(2.0)  # 少し待ってリトライ
                    continue

                # 2. 画像形式変換 (PIL -> OpenCV)
                current_image_pil = screenshot.convert("RGB")
                current_image_cv = np.array(current_image_pil)
                current_image_cv = cv2.cvtColor(current_image_cv, cv2.COLOR_RGB2BGR)

                if current_image_cv is None or current_image_cv.size == 0:
                    logger.error(
                        "エラー: 画像データの変換に失敗しました (Noneまたはサイズ0)。"
                    )
                    self.error_occurred_in_screenshot_thread = (
                        True  # error_occurred_in_thread から変更
                    )
                    time.sleep(2.0)
                    continue

                # 3. 前回の画像と比較
                if self.last_screenshot_image is None:  # last_image から変更
                    logger.info(
                        "最初のスクリーンショット画像を取得しました。保存します。"
                    )
                    self.save_screenshot_image(current_image_cv)  # save_image から変更
                    self.last_screenshot_image = current_image_cv  # last_image から変更
                else:
                    if not self.is_similar(
                        current_image_cv, self.last_screenshot_image
                    ):  # last_image から変更
                        logger.info(
                            "新しいスクリーンショット画像または類似していない画像を検出しました。保存します。"
                        )
                        self.save_screenshot_image(
                            current_image_cv
                        )  # save_image から変更
                        self.last_screenshot_image = (
                            current_image_cv  # last_image から変更
                        )
                    else:
                        # logger.debug("類似画像のためスキップ")
                        pass

            except (OSError, UnidentifiedImageError) as e:
                logger.exception(
                    f"エラー (スクショキャプチャ/変換): {type(e).__name__} - {e}"
                )
                self.error_occurred_in_screenshot_thread = (
                    True  # error_occurred_in_thread から変更
                )
                time.sleep(5.0)
            except cv2.error as e:
                logger.exception(
                    f"エラー (スクショ - OpenCV): {type(e).__name__} - {e}"
                )
                self.error_occurred_in_screenshot_thread = (
                    True  # error_occurred_in_thread から変更
                )
                time.sleep(5.0)
            except Exception as e:
                logger.exception(
                    f"エラー (スクショループ): 予期せぬエラーが発生しました - {type(e).__name__}: {e}"
                )
                self.error_occurred_in_screenshot_thread = (
                    True  # error_occurred_in_thread から変更
                )
                time.sleep(5.0)

            # 次のキャプチャまでの待機時間
            wait_start_time = time.time()
            while (
                self.is_capturing_screenshot and time.time() - wait_start_time < 2.0
            ):  # is_capturing から変更
                time.sleep(0.1)

        logger.info("スクリーンショットキャプチャループが終了しました。")

    def is_similar(self, img1_cv, img2_cv, threshold=0.95):
        """2つの画像の類似度を計算する (差分ベースの簡易比較)"""
        try:
            gray1 = cv2.cvtColor(img1_cv, cv2.COLOR_BGR2GRAY)
            gray2 = cv2.cvtColor(img2_cv, cv2.COLOR_BGR2GRAY)
            h1, w1 = gray1.shape
            h2, w2 = gray2.shape

            if h1 != h2 or w1 != w2:
                # logger.debug(f"画像サイズが異なるためリサイズします: ({w1}x{h1}) vs ({w2}x{h2})")
                if h1 * w1 < h2 * w2:
                    gray2 = cv2.resize(gray2, (w1, h1), interpolation=cv2.INTER_AREA)
                else:
                    gray1 = cv2.resize(gray1, (w2, h2), interpolation=cv2.INTER_AREA)

            diff = cv2.absdiff(gray1, gray2)
            non_zero_count = np.count_nonzero(diff)
            total_pixels = gray1.size
            if total_pixels == 0:
                logger.warning("警告: 類似度計算中の画像サイズが0です。")
                return False

            similarity = 1.0 - (non_zero_count / total_pixels)
            # logger.debug(f"類似度: {similarity:.4f}")
            return similarity >= threshold

        except cv2.error as e:
            logger.exception(f"エラー (類似度計算 - OpenCV): {type(e).__name__} - {e}")
            self.error_occurred_in_screenshot_thread = (
                True  # error_occurred_in_thread から変更
            )
            return False
        except Exception as e:
            logger.exception(
                f"エラー (類似度計算): 予期せぬエラー - {type(e).__name__}: {e}"
            )
            self.error_occurred_in_screenshot_thread = (
                True  # error_occurred_in_thread から変更
            )
            return False

    def save_screenshot_image(self, image_cv):  # save_image から変更
        """スクリーンショット画像をファイルに保存する。エラー発生時はログに記録。"""
        save_path = None
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
            filename = f"screenshot_{timestamp}.png"
            save_path = os.path.join(self.capture_save_path, filename)

            if image_cv is None or image_cv.size == 0:
                logger.warning(
                    f"警告: 保存しようとしたスクリーンショット画像データが無効です (Noneまたはサイズ0)。パス: {save_path}"
                )
                self.error_occurred_in_screenshot_thread = (
                    True  # error_occurred_in_thread から変更
                )
                return

            success = cv2.imwrite(save_path, image_cv, [cv2.IMWRITE_PNG_COMPRESSION, 3])

            if success:
                self.screenshot_saved_count += 1  # saved_count から変更
                self.last_saved_screenshot_filename = os.path.basename(
                    save_path
                )  # last_saved_filename から変更
                logger.info(f"スクリーンショット画像を保存しました: {save_path}")
            else:
                # imwriteがFalseを返した場合
                error_msg = (
                    f"エラー (保存): cv2.imwrite が False を返しました。ファイル書き込みに失敗した可能性があります。\n"
                    f" - 保存試行パス: {save_path}\n"
                    f" - 画像サイズ: {image_cv.shape if image_cv is not None else 'None'}\n"
                    f" - 考えられる原因: 書き込み権限不足、ディスク容量不足、パス名の問題、画像データ破損など"
                )
                logger.error(error_msg)
                self.error_occurred_in_screenshot_thread = (
                    True  # error_occurred_in_thread から変更
                )

        except cv2.error as e:
            logger.exception(
                f"エラー (スクショ保存 - OpenCV): {type(e).__name__} - {e}\n - 保存試行パス: {save_path}"
            )
            self.error_occurred_in_screenshot_thread = (
                True  # error_occurred_in_thread から変更
            )
        except OSError as e:
            logger.exception(
                f"エラー (スクショ保存 - OS): {type(e).__name__} - {e}\n - 保存試行パス: {save_path}"
            )
            self.error_occurred_in_screenshot_thread = (
                True  # error_occurred_in_thread から変更
            )
        except Exception as e:
            logger.exception(
                f"エラー (スクショ保存): 予期せぬエラー - {type(e).__name__}: {e}\n - 保存試行パス: {save_path}"
            )
            self.error_occurred_in_screenshot_thread = (
                True  # error_occurred_in_thread から変更
            )

    def on_closing(self):
        """ウィンドウが閉じられたときの処理"""
        if self.is_capturing_screenshot or self.is_recording:  # is_capturing から変更
            tasks_running = []
            if self.is_capturing_screenshot:
                tasks_running.append("スクリーンショットキャプチャ")
            if self.is_recording:
                tasks_running.append("録画")
            running_tasks_str = " と ".join(tasks_running)

            if messagebox.askokcancel(
                "確認", f"{running_tasks_str}が実行中です。\n本当に終了しますか？"
            ):
                logger.info("ユーザー操作により終了します...")
                self.stop_all_tasks()  # stop_capture から変更
                self.root.destroy()
                logger.info("アプリケーションを終了しました。")
            else:
                logger.info("終了操作がキャンセルされました。")
                return
        else:
            logger.info("アプリケーションを終了しました。")
            self.root.destroy()


if __name__ == "__main__":
    # Tkinterのルートウィンドウを作成
    root = tk.Tk()
    # アプリケーションクラスのインスタンスを作成
    app = SlideCaptureApp(root)
    # Tkinterのイベントループを開始
    root.mainloop()
