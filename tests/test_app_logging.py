# -*- coding: utf-8 -*-
"""起動時ロガー設定の回帰テスト"""

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).parent.parent))

from app_logging import configure_initial_logging


class TestInitialLoggingConfiguration(unittest.TestCase):
    """developer_mode に応じた初期ログハンドラ設定を検証する"""

    def test_configure_initial_logging_omits_stream_handler_when_developer_mode_disabled(
        self,
    ):
        """developer_mode=False では起動時にコンソール出力しない"""
        file_handler = object()

        with (
            patch("app_logging.logging.FileHandler", return_value=file_handler) as file_cls,
            patch("app_logging.logging.StreamHandler") as stream_cls,
            patch("app_logging.logging.basicConfig") as basic_config,
        ):
            configure_initial_logging(
                {
                    "level": "INFO",
                    "filename": "app.log",
                    "format": "%(levelname)s:%(message)s",
                },
                developer_mode=False,
            )

        file_cls.assert_called_once_with("app.log", encoding="utf-8")
        stream_cls.assert_not_called()
        basic_config.assert_called_once()
        self.assertEqual([file_handler], basic_config.call_args.kwargs["handlers"])

    def test_configure_initial_logging_adds_stream_handler_when_developer_mode_enabled(
        self,
    ):
        """developer_mode=True では起動時にコンソール出力も有効化する"""
        file_handler = object()
        stream_handler = object()

        with (
            patch("app_logging.logging.FileHandler", return_value=file_handler) as file_cls,
            patch(
                "app_logging.logging.StreamHandler", return_value=stream_handler
            ) as stream_cls,
            patch("app_logging.logging.basicConfig") as basic_config,
        ):
            configure_initial_logging(
                {
                    "level": "DEBUG",
                    "filename": "debug.log",
                    "format": "%(levelname)s:%(message)s",
                },
                developer_mode=True,
            )

        file_cls.assert_called_once_with("debug.log", encoding="utf-8")
        stream_cls.assert_called_once()
        basic_config.assert_called_once()
        self.assertEqual(
            [file_handler, stream_handler],
            basic_config.call_args.kwargs["handlers"],
        )


if __name__ == "__main__":
    unittest.main()
