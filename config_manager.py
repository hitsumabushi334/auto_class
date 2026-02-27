# -*- coding: utf-8 -*-
"""
Configuration management module for auto_class application.
Handles loading and validation of application configuration.

起動時に設定ファイルが存在しない、あるいは不正な場合は
DEFAULT_CONFIG を使って自動的にファイルを生成してから読み込む。
"""

import copy
import json
import os
import sys
import logging
from typing import Dict, Any, List, Optional

logger = logging.getLogger(__name__)

# --------------------------------------------------------------------------- #
# デフォルト設定 (新方式: api.models にオブジェクトリストを使用)
# --------------------------------------------------------------------------- #
DEFAULT_CONFIG: Dict[str, Any] = {
    "api": {
        "gemini_api_key": "",
        "models": [
            {
                "model_name": "gemini-3-flash-preview",
                "description": "一般用途、無料枠",
            },
            {
                "model_name": "gemini-3.1-pro-preview",
                "description": "高性能モデル、有料枠のみ使用可",
            },
        ],
        "default_model_index": 0,
    },
    "audio": {
        "no_sound_timeout_seconds": 180,
        "silence_threshold": 0.01,
    },
    "screenshot": {
        "similarity_threshold": 0.83,
        "diff_pixel_threshold": 10,
    },
    "ui": {
        "window_width": 600,
        "window_height": 550,
        "min_width": 600,
        "min_height": 550,
    },
    "logging": {
        "level": "INFO",
        "filename": "slide_capture_app.log",
        "format": "%(asctime)s - %(levelname)s - %(threadName)s - %(message)s",
    },
}


def _is_valid_config(data: Any) -> bool:
    """設定データが新方式として有効かを検証する。

    必須条件:
    - dict 型
    - "api" キーが存在
    - "api.models" が list[dict] 型（model_name キーを持つ）
    """
    if not isinstance(data, dict):
        return False
    api = data.get("api")
    if not isinstance(api, dict):
        return False
    models = api.get("models")
    if not isinstance(models, list) or len(models) == 0:
        return False
    # 全要素が dict で model_name を持つ場合のみ有効
    return all(isinstance(m, dict) and "model_name" in m for m in models)


class ConfigManager:
    """アプリケーション設定の読み込みと検証を管理するクラス。"""

    def __init__(self, config_path: Optional[str] = None):
        """ConfigManager を初期化する。

        Args:
            config_path: 設定ファイルのパス。None または空の場合は自動探索する。
        """
        self.config_path = self._resolve_config_path(config_path)
        self._config: Dict[str, Any] = {}
        self.load_config()

    # ----------------------------------------------------------------------- #
    # パス解決
    # ----------------------------------------------------------------------- #

    def _resolve_config_path(self, config_path: Optional[str]) -> str:
        """config.json のパスを解決する。

        優先順位:
        1. 明示的に指定されたパス
        2. PyInstaller でビルドされた場合は exe と同じディレクトリ
        3. このモジュールと同じディレクトリ
        4. カレントワーキングディレクトリ
        """
        # 1. 明示的に指定されたパス
        if config_path:
            return os.path.abspath(config_path)

        # 2. exe と同じディレクトリ (frozen)
        if getattr(sys, "frozen", False):
            return os.path.join(os.path.dirname(sys.executable), "config.json")

        # 3. モジュールと同じディレクトリ
        module_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "config.json"
        )
        if os.path.isfile(module_path):
            return module_path

        # 4. カレントワーキングディレクトリ
        cwd_path = os.path.join(os.getcwd(), "config.json")
        if os.path.isfile(cwd_path):
            return cwd_path

        # どこにも存在しない場合はモジュールディレクトリの config.json を返す
        if getattr(sys, "frozen", False):
            return os.path.join(os.path.dirname(sys.executable), "config.json")
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

    # ----------------------------------------------------------------------- #
    # 設定ファイルの書き込み (デフォルト値で生成/上書き)
    # ----------------------------------------------------------------------- #

    def _write_default_config(self) -> bool:
        """DEFAULT_CONFIG を config_path に書き込む。

        Returns:
            書き込み成功時 True、PermissionError 等の失敗時は False。
        """
        try:
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
        except (OSError, ValueError):
            pass  # ルートディレクトリなど makedirs が不要なケース

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
            logger.info(f"デフォルト設定ファイルを生成しました: {self.config_path}")
            return True
        except PermissionError as e:
            logger.error(
                f"設定ファイルへの書き込み権限がありません。インメモリのデフォルト設定で起動します。"
                f" path={self.config_path}, error={e}"
            )
            return False
        except OSError as e:
            logger.error(
                f"設定ファイルの書き込みに失敗しました。インメモリのデフォルト設定で起動します。"
                f" path={self.config_path}, error={e}"
            )
            return False

    # ----------------------------------------------------------------------- #
    # 設定の読み込み
    # ----------------------------------------------------------------------- #

    def load_config(self) -> None:
        """設定ファイルを読み込む。

        ファイルが存在しない、パースエラー、または不正なフォーマット（旧形式を含む）の場合は
        DEFAULT_CONFIG でファイルを上書き生成してから読み込む。
        """
        needs_overwrite = False
        reason = ""

        # ファイルが存在しない場合
        if not os.path.exists(self.config_path):
            reason = f"設定ファイルが見つかりません (path={self.config_path})"
            needs_overwrite = True
        else:
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if not _is_valid_config(data):
                    reason = (
                        f"設定ファイルのフォーマットが不正です（旧形式または必須キー不足）。"
                        f" path={self.config_path}"
                    )
                    needs_overwrite = True
                else:
                    self._config = data
                    logger.info(f"設定ファイルを読み込みました: {self.config_path}")
                    return
            except json.JSONDecodeError as e:
                reason = f"設定ファイルの JSON パースに失敗しました。 error={e}"
                needs_overwrite = True

        if needs_overwrite:
            logger.warning(f"{reason} — デフォルト設定でファイルを上書き生成します。")
            wrote = self._write_default_config()
            if wrote:
                # 書き込んだファイルを読み込む
                try:
                    with open(self.config_path, "r", encoding="utf-8") as f:
                        self._config = json.load(f)
                    return
                except Exception as e:
                    logger.error(f"上書きした設定ファイルの読み込みに失敗しました: {e}")

            # 書き込み不可などでファイルが読めない場合はインメモリで対応
            logger.warning("インメモリのデフォルト設定を使用します。")
            self._config = copy.deepcopy(DEFAULT_CONFIG)

    # ----------------------------------------------------------------------- #
    # 汎用 getter
    # ----------------------------------------------------------------------- #

    def get(self, key: str, default: Any = None) -> Any:
        """ドット区切りのキーで設定値を取得する。

        Args:
            key: ドット区切りのキー (例: 'api.gemini_api_key')
            default: キーが存在しない場合のデフォルト値

        Returns:
            設定値
        """
        keys = key.split(".")
        value = self._config

        try:
            for k in keys:
                value = value[k]
            return value
        except (KeyError, TypeError):
            if default is not None:
                return default
            raise KeyError(f"設定キーが見つかりません: {key}")

    # ----------------------------------------------------------------------- #
    # ドメイン固有 getter
    # ----------------------------------------------------------------------- #

    def get_api_key(self) -> str:
        """APIキーを取得する（設定ファイルの値のみを使用）。"""
        return self.get("api.gemini_api_key", "")

    def get_model_options(self) -> List[str]:
        """利用可能なモデル名の一覧を返す。"""
        models = self.get("api.models", [])
        return [
            m["model_name"] for m in models if isinstance(m, dict) and "model_name" in m
        ]

    def get_default_model(self) -> str:
        """デフォルトモデル名を返す。"""
        models = self.get_model_options()
        if not models:
            raise ValueError("モデルが設定されていません")
        default_index = self.get("api.default_model_index", 0)
        if default_index >= len(models):
            logger.warning(
                f"default_model_index ({default_index}) がモデル数 ({len(models)}) を超えています。0 を使用します。"
            )
            default_index = 0
        return models[default_index]

    def get_model_description(self, model_name: str) -> str:
        """指定したモデルの説明文を返す。"""
        models = self.get("api.models", [])
        for m in models:
            if isinstance(m, dict) and m.get("model_name") == model_name:
                return m.get("description", "用途: 不明")
        return "用途: 不明"

    def get_audio_settings(self) -> Dict[str, Any]:
        """音声設定を返す。"""
        return {
            "no_sound_timeout_seconds": self.get("audio.no_sound_timeout_seconds", 180),
            "silence_threshold": self.get("audio.silence_threshold", 0.01),
        }

    def get_screenshot_settings(self) -> Dict[str, Any]:
        """スクリーンショット設定を返す。"""
        return {
            "similarity_threshold": self.get("screenshot.similarity_threshold", 0.83),
            "diff_pixel_threshold": self.get("screenshot.diff_pixel_threshold", 10),
        }

    def get_ui_settings(self) -> Dict[str, Any]:
        """UI設定を返す。"""
        return {
            "window_width": self.get("ui.window_width", 600),
            "window_height": self.get("ui.window_height", 550),
            "min_width": self.get("ui.min_width", 600),
            "min_height": self.get("ui.min_height", 550),
        }

    def get_logging_settings(self) -> Dict[str, Any]:
        """ロギング設定を返す。"""
        return {
            "level": self.get("logging.level", "INFO"),
            "filename": self.get("logging.filename", "slide_capture_app.log"),
            "format": self.get(
                "logging.format",
                "%(asctime)s - %(levelname)s - %(threadName)s - %(message)s",
            ),
        }


# --------------------------------------------------------------------------- #
# グローバルインスタンス管理
# --------------------------------------------------------------------------- #

_config_manager: Optional[ConfigManager] = None


def get_config_manager() -> ConfigManager:
    """グローバルな ConfigManager インスタンスを返す。"""
    global _config_manager
    if _config_manager is None:
        _config_manager = ConfigManager()
    return _config_manager


def reload_config() -> None:
    """設定ファイルを再読み込みする。"""
    global _config_manager
    if _config_manager is not None:
        _config_manager.load_config()
