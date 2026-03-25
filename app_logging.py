# -*- coding: utf-8 -*-
"""アプリケーションのロガー初期化補助"""

import logging
import sys


def configure_initial_logging(logging_settings: dict, developer_mode: bool) -> None:
    """起動直後の root logger を設定する。"""
    handlers = [
        logging.FileHandler(logging_settings["filename"], encoding="utf-8"),
    ]
    if developer_mode:
        handlers.append(logging.StreamHandler(sys.stdout))

    logging.basicConfig(
        level=getattr(logging, logging_settings["level"].upper()),
        format=logging_settings["format"],
        handlers=handlers,
        force=True,
    )
