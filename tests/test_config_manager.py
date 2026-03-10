# -*- coding: utf-8 -*-
"""config_manager.py のユニットテスト (TDD)"""

import json
import os
import sys
import stat
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

# プロジェクトルートをパスに追加
sys.path.insert(0, str(Path(__file__).parent.parent))

from config_manager import ConfigManager, DEFAULT_CONFIG


class TestConfigManagerDefaultGeneration(unittest.TestCase):
    """config.json が存在しない場合のデフォルト生成テスト"""

    def test_creates_config_file_when_missing(self):
        """config.json が存在しない場合、自動的に生成されること"""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = os.path.join(tmpdir, "config.json")
            manager = ConfigManager(config_path=config_path)
            self.assertTrue(os.path.exists(config_path))

    def test_flag_set_when_file_is_missing(self):
        """config.json が存在しない場合、config_was_auto_generated フラグが True になること"""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = os.path.join(tmpdir, "config.json")
            manager = ConfigManager(config_path=config_path)
            self.assertTrue(manager.config_was_auto_generated)
            self.assertNotEqual(manager.config_auto_generated_reason, "")

    def test_generated_config_has_new_models_format(self):
        """生成された config.json が新方式の models 形式（リスト of dict）を持つこと"""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = os.path.join(tmpdir, "config.json")
            ConfigManager(config_path=config_path)
            with open(config_path, encoding="utf-8") as f:
                data = json.load(f)
            models = data["api"]["models"]
            self.assertIsInstance(models, list)
            self.assertGreater(len(models), 0)
            # 各要素が dict で model_name と description を持つこと
            for model in models:
                self.assertIsInstance(model, dict)
                self.assertIn("model_name", model)
                self.assertIn("description", model)

    def test_generated_config_no_bom_and_utf8(self):
        """生成された config.json が UTF-8 で保存されており、日本語が化けないこと"""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = os.path.join(tmpdir, "config.json")
            ConfigManager(config_path=config_path)
            with open(config_path, encoding="utf-8") as f:
                content = f.read()
            # デフォルト設定に日本語の説明があれば文字化けしないこと
            self.assertNotIn("\\u", content)  # ensure_ascii=False なら \\u は不要


class TestConfigManagerOverwrite(unittest.TestCase):
    """壊れた/旧形式の config.json をデフォルトで上書きするテスト"""

    def test_overwrites_broken_json_and_logs_error(self):
        """壊れた JSON ファイルが存在する場合、上書きされてデフォルト値が読み込まれること"""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = os.path.join(tmpdir, "config.json")
            # 壊れた JSON を書き込む
            with open(config_path, "w", encoding="utf-8") as f:
                f.write("{ this is not valid json }")
            manager = ConfigManager(config_path=config_path)
            # ロード後に正常なデフォルト値が得られること
            models = manager.get_model_options()
            self.assertIsInstance(models, list)
            self.assertGreater(len(models), 0)

    def test_overwrites_old_format_config(self):
        """旧形式（models が文字列リスト）が存在する場合、新形式で上書きされること"""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = os.path.join(tmpdir, "config.json")
            old_config = {
                "api": {
                    "gemini_api_key": "",
                    "models": ["gemini-old-model"],  # 旧形式（文字列リスト）
                    "default_model_index": 0,
                    "model_descriptions": {"gemini-old-model": "旧説明文"},
                }
            }
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(old_config, f, ensure_ascii=False)
            manager = ConfigManager(config_path=config_path)
            # 上書き後はデフォルトモデルの一覧（新形式）が返ること
            models = manager.get_model_options()
            self.assertIsInstance(models, list)
            for name in models:
                self.assertIsInstance(name, str)

    def test_permission_error_falls_back_to_in_memory(self):
        """書き込み権限がない場合でもクラッシュせず、インメモリのデフォルト設定で起動できること"""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = os.path.join(tmpdir, "config.json")
            # PermissionError を模擬
            with patch("builtins.open", side_effect=PermissionError("no write access")):
                # 例外が投げられず起動できること
                try:
                    manager = ConfigManager(config_path=config_path)
                    models = manager.get_model_options()
                    self.assertIsInstance(models, list)
                except PermissionError:
                    self.fail("PermissionError が捕捉されずに伝播しました")


class TestConfigManagerGetters(unittest.TestCase):
    """新方式の設定ファイルを使った getter のテスト"""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.config_path = os.path.join(self.tmpdir, "config.json")
        new_config = {
            "api": {
                "gemini_api_key": "test-api-key",
                "models": [
                    {"model_name": "model-alpha", "description": "説明A"},
                    {"model_name": "model-beta", "description": "説明B"},
                ],
                "default_model_index": 0,
            },
            "audio": {"no_sound_timeout_seconds": 180, "silence_threshold": 0.01},
            "screenshot": {"similarity_threshold": 0.83, "diff_pixel_threshold": 10},
            "ui": {
                "window_width": 600,
                "window_height": 550,
                "min_width": 600,
                "min_height": 550,
            },
            "logging": {
                "level": "INFO",
                "filename": "app.log",
                "format": "%(asctime)s - %(message)s",
            },
        }
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(new_config, f, ensure_ascii=False)
        self.manager = ConfigManager(config_path=self.config_path)

    def tearDown(self):
        import shutil

        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_get_model_options_returns_name_list(self):
        """get_model_options が model_name の文字列リストを返すこと"""
        options = self.manager.get_model_options()
        self.assertEqual(options, ["model-alpha", "model-beta"])

    def test_get_default_model_returns_first(self):
        """get_default_model がインデックス 0 のモデル名を返すこと"""
        self.assertEqual(self.manager.get_default_model(), "model-alpha")

    def test_get_model_description_returns_correct_description(self):
        """get_model_description が対応する description を返すこと"""
        self.assertEqual(self.manager.get_model_description("model-alpha"), "説明A")
        self.assertEqual(self.manager.get_model_description("model-beta"), "説明B")

    def test_get_model_description_unknown_returns_fallback(self):
        """未知のモデル名に対しては「用途: 不明」を返すこと"""
        self.assertEqual(
            self.manager.get_model_description("unknown-model"), "用途: 不明"
        )

    def test_get_api_key_reads_from_json_not_env(self):
        """get_api_key が環境変数を無視して config.json の値を返すこと"""
        with patch.dict(os.environ, {"GEMINI_API_KEY": "env-key-should-not-be-used"}):
            key = self.manager.get_api_key()
            self.assertEqual(key, "test-api-key")


# ===================================================================
# 新機能のテスト: update_config / reset_to_default / get_app_settings
# ===================================================================


class TestConfigManagerUpdateConfig(unittest.TestCase):
    """update_config メソッドのテスト"""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.config_path = os.path.join(self.tmpdir, "config.json")
        # 正常な設定ファイルを用意
        initial_config = {
            "api": {
                "gemini_api_key": "initial-key",
                "models": [{"model_name": "model-a", "description": "説明A"}],
                "default_model_index": 0,
            },
            "audio": {"no_sound_timeout_seconds": 180, "silence_threshold": 0.01},
            "screenshot": {"similarity_threshold": 0.83, "diff_pixel_threshold": 10},
            "ui": {"window_width": 600, "window_height": 550, "min_width": 600, "min_height": 550},
            "logging": {"level": "INFO", "filename": "app.log", "format": "%(asctime)s"},
            "app": {"note_output_mode": "MD", "developer_mode": False},
        }
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(initial_config, f, ensure_ascii=False)
        self.manager = ConfigManager(config_path=self.config_path)

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_update_config_updates_in_memory_state(self):
        """update_config 呼び出し後、get_api_key が新しいキーを返すこと"""
        new_config = {
            "api": {
                "gemini_api_key": "new-key",
                "models": [{"model_name": "model-a", "description": "説明A"}],
                "default_model_index": 0,
            },
            "audio": {"no_sound_timeout_seconds": 180, "silence_threshold": 0.01},
            "screenshot": {"similarity_threshold": 0.83, "diff_pixel_threshold": 10},
            "ui": {"window_width": 600, "window_height": 550, "min_width": 600, "min_height": 550},
            "logging": {"level": "INFO", "filename": "app.log", "format": "%(asctime)s"},
            "app": {"note_output_mode": "MD", "developer_mode": False},
        }
        self.manager.update_config(new_config)
        self.assertEqual(self.manager.get_api_key(), "new-key")

    def test_update_config_writes_to_file(self):
        """update_config 呼び出し後、config.json が新しい値で上書きされること"""
        new_config = {
            "api": {
                "gemini_api_key": "file-check-key",
                "models": [{"model_name": "model-a", "description": "説明A"}],
                "default_model_index": 0,
            },
            "audio": {"no_sound_timeout_seconds": 180, "silence_threshold": 0.01},
            "screenshot": {"similarity_threshold": 0.83, "diff_pixel_threshold": 10},
            "ui": {"window_width": 600, "window_height": 550, "min_width": 600, "min_height": 550},
            "logging": {"level": "INFO", "filename": "app.log", "format": "%(asctime)s"},
            "app": {"note_output_mode": "MD", "developer_mode": False},
        }
        self.manager.update_config(new_config)
        with open(self.config_path, encoding="utf-8") as f:
            saved = json.load(f)
        self.assertEqual(saved["api"]["gemini_api_key"], "file-check-key")

    def test_update_config_subsequent_get_returns_new_value(self):
        """update_config 後、get() でも新しい値が返ること（キャッシュ不整合がないこと）"""
        new_config = {
            "api": {
                "gemini_api_key": "another-key",
                "models": [{"model_name": "model-a", "description": "説明A"}],
                "default_model_index": 0,
            },
            "audio": {"no_sound_timeout_seconds": 300, "silence_threshold": 0.02},
            "screenshot": {"similarity_threshold": 0.9, "diff_pixel_threshold": 5},
            "ui": {"window_width": 800, "window_height": 600, "min_width": 800, "min_height": 600},
            "logging": {"level": "DEBUG", "filename": "debug.log", "format": "%(asctime)s"},
            "app": {"note_output_mode": "Word", "developer_mode": True},
        }
        self.manager.update_config(new_config)
        self.assertEqual(self.manager.get("audio.no_sound_timeout_seconds"), 300)
        self.assertEqual(self.manager.get("app.note_output_mode"), "Word")


class TestConfigManagerResetToDefault(unittest.TestCase):
    """reset_to_default メソッドのテスト"""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.config_path = os.path.join(self.tmpdir, "config.json")
        # カスタム設定を書き込む
        custom_config = {
            "api": {
                "gemini_api_key": "custom-key",
                "models": [{"model_name": "custom-model", "description": "カスタムモデル"}],
                "default_model_index": 0,
            },
            "audio": {"no_sound_timeout_seconds": 9999, "silence_threshold": 0.99},
            "screenshot": {"similarity_threshold": 0.5, "diff_pixel_threshold": 100},
            "ui": {"window_width": 1200, "window_height": 900, "min_width": 1200, "min_height": 900},
            "logging": {"level": "DEBUG", "filename": "custom.log", "format": "%(message)s"},
            "app": {"note_output_mode": "Word", "developer_mode": True},
        }
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(custom_config, f, ensure_ascii=False)
        self.manager = ConfigManager(config_path=self.config_path)

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_reset_to_default_clears_in_memory_state(self):
        """reset_to_default 後、audio 設定がデフォルト値になること"""
        self.manager.reset_to_default()
        audio = self.manager.get_audio_settings()
        self.assertEqual(audio["no_sound_timeout_seconds"], 180)

    def test_reset_to_default_writes_default_to_file(self):
        """reset_to_default 後、config.json の api_key がデフォルト（空文字）になること"""
        self.manager.reset_to_default()
        with open(self.config_path, encoding="utf-8") as f:
            saved = json.load(f)
        self.assertEqual(saved["api"]["gemini_api_key"], "")

    def test_reset_to_default_restores_app_section_defaults(self):
        """reset_to_default 後、app.note_output_mode が 'MD' に戻ること"""
        self.manager.reset_to_default()
        mode = self.manager.get("app.note_output_mode")
        self.assertEqual(mode, "MD")


class TestConfigManagerGetCurrentConfig(unittest.TestCase):
    """get_current_config メソッドのテスト"""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.config_path = os.path.join(self.tmpdir, "config.json")
        cfg = {
            "api": {
                "gemini_api_key": "test",
                "models": [{"model_name": "m", "description": "d"}],
                "default_model_index": 0,
            },
            "audio": {"no_sound_timeout_seconds": 180, "silence_threshold": 0.01},
            "screenshot": {"similarity_threshold": 0.83, "diff_pixel_threshold": 10},
            "ui": {"window_width": 600, "window_height": 550, "min_width": 600, "min_height": 550},
            "logging": {"level": "INFO", "filename": "app.log", "format": "%(asctime)s"},
            "app": {"note_output_mode": "MD", "developer_mode": False},
        }
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False)
        self.manager = ConfigManager(config_path=self.config_path)

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_get_current_config_returns_dict_copy(self):
        """get_current_config が辞書を返し、変更してもインスタンスに影響しないこと"""
        cfg = self.manager.get_current_config()
        self.assertIsInstance(cfg, dict)
        cfg["api"]["gemini_api_key"] = "mutated"
        # インスタンスの内部状態が変わっていないこと
        self.assertEqual(self.manager.get_api_key(), "test")


class TestConfigManagerAppSettings(unittest.TestCase):
    """app セクション (note_output_mode / developer_mode) のテスト"""

    def _make_manager(self, app_section):
        tmpdir = tempfile.mkdtemp()
        config_path = os.path.join(tmpdir, "config.json")
        cfg = {
            "api": {
                "gemini_api_key": "",
                "models": [{"model_name": "m", "description": "d"}],
                "default_model_index": 0,
            },
            "audio": {"no_sound_timeout_seconds": 180, "silence_threshold": 0.01},
            "screenshot": {"similarity_threshold": 0.83, "diff_pixel_threshold": 10},
            "ui": {"window_width": 600, "window_height": 550, "min_width": 600, "min_height": 550},
            "logging": {"level": "INFO", "filename": "app.log", "format": "%(asctime)s"},
            "app": app_section,
        }
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False)
        return ConfigManager(config_path=config_path), tmpdir

    def test_get_app_settings_returns_md_mode_by_default(self):
        """DEFAULT_CONFIG の app.note_output_mode が 'MD' であること"""
        manager, tmpdir = self._make_manager({"note_output_mode": "MD", "developer_mode": False})
        try:
            settings = manager.get_app_settings()
            self.assertEqual(settings["note_output_mode"], "MD")
        finally:
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_get_app_settings_returns_word_mode_when_set(self):
        """app.note_output_mode が 'Word' のとき get_app_settings で 'Word' が返ること"""
        manager, tmpdir = self._make_manager({"note_output_mode": "Word", "developer_mode": False})
        try:
            settings = manager.get_app_settings()
            self.assertEqual(settings["note_output_mode"], "Word")
        finally:
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_get_app_settings_returns_developer_mode_false_by_default(self):
        """app.developer_mode が False であること"""
        manager, tmpdir = self._make_manager({"note_output_mode": "MD", "developer_mode": False})
        try:
            settings = manager.get_app_settings()
            self.assertFalse(settings["developer_mode"])
        finally:
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_default_config_has_app_section(self):
        """DEFAULT_CONFIG に 'app' セクションが含まれること"""
        from config_manager import DEFAULT_CONFIG
        self.assertIn("app", DEFAULT_CONFIG)
        self.assertIn("note_output_mode", DEFAULT_CONFIG["app"])
        self.assertIn("developer_mode", DEFAULT_CONFIG["app"])

    def test_default_config_app_note_output_mode_is_md(self):
        """DEFAULT_CONFIG の app.note_output_mode が 'MD' であること"""
        from config_manager import DEFAULT_CONFIG
        self.assertEqual(DEFAULT_CONFIG["app"]["note_output_mode"], "MD")

    def test_default_config_app_developer_mode_is_false(self):
        """DEFAULT_CONFIG の app.developer_mode が False であること"""
        from config_manager import DEFAULT_CONFIG
        self.assertFalse(DEFAULT_CONFIG["app"]["developer_mode"])


if __name__ == "__main__":
    unittest.main()
