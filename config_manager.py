# -*- coding: utf-8 -*-
"""
Configuration management module for auto_class application.
Handles loading and validation of application configuration.
"""

import json
import os
import logging
from typing import Dict, Any, List

logger = logging.getLogger(__name__)


class ConfigManager:
    """Manages application configuration loading and validation."""
    
    def __init__(self, config_path: str = None):
        """
        Initialize configuration manager.
        
        Args:
            config_path (str): Path to config file. If None, uses default 'config.json'
        """
        if config_path is None:
            # Default to config.json in the same directory as this module
            current_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(current_dir, "config.json")
        
        self.config_path = config_path
        self._config = {}
        self.load_config()
    
    def load_config(self) -> None:
        """Load configuration from file."""
        try:
            if not os.path.exists(self.config_path):
                raise FileNotFoundError(f"Configuration file not found: {self.config_path}")
            
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self._config = json.load(f)
            
            logger.info(f"Configuration loaded from: {self.config_path}")
            
        except Exception as e:
            logger.error(f"Failed to load configuration: {e}")
            raise
    
    def get(self, key: str, default: Any = None) -> Any:
        """
        Get configuration value by dot-notation key.
        
        Args:
            key (str): Configuration key in dot notation (e.g., 'api.gemini_api_key')
            default (Any): Default value if key is not found
            
        Returns:
            Any: Configuration value
        """
        keys = key.split('.')
        value = self._config
        
        try:
            for k in keys:
                value = value[k]
            return value
        except (KeyError, TypeError):
            if default is not None:
                return default
            raise KeyError(f"Configuration key not found: {key}")
    
    def get_api_key(self) -> str:
        """Get API key, preferring environment variable over config file."""
        # First try environment variable
        env_key = os.environ.get("GEMINI_API_KEY")
        if env_key:
            return env_key
        
        # Fall back to config file
        return self.get("api.gemini_api_key")
    
    def get_model_options(self) -> List[str]:
        """Get available Gemini model options."""
        return self.get("api.models", [])
    
    def get_default_model(self) -> str:
        """Get default Gemini model."""
        models = self.get_model_options()
        default_index = self.get("api.default_model_index", 0)
        
        if not models:
            raise ValueError("No models configured")
        
        if default_index >= len(models):
            logger.warning(f"Default model index {default_index} out of range, using 0")
            default_index = 0
        
        return models[default_index]
    
    def get_model_description(self, model_name: str) -> str:
        """Get description for a specific model."""
        descriptions = self.get("api.model_descriptions", {})
        return descriptions.get(model_name, "用途: 不明")
    
    def get_audio_settings(self) -> Dict[str, Any]:
        """Get audio configuration settings."""
        return {
            'no_sound_timeout_seconds': self.get('audio.no_sound_timeout_seconds', 180),
            'silence_threshold': self.get('audio.silence_threshold', 0.01)
        }
    
    def get_ui_settings(self) -> Dict[str, Any]:
        """Get UI configuration settings."""
        return {
            'window_width': self.get('ui.window_width', 600),
            'window_height': self.get('ui.window_height', 550),
            'min_width': self.get('ui.min_width', 600),
            'min_height': self.get('ui.min_height', 550)
        }
    
    def get_logging_settings(self) -> Dict[str, Any]:
        """Get logging configuration settings."""
        return {
            'level': self.get('logging.level', 'INFO'),
            'filename': self.get('logging.filename', 'slide_capture_app.log'),
            'format': self.get('logging.format', '%(asctime)s - %(levelname)s - %(threadName)s - %(message)s')
        }


# Global configuration instance
_config_manager = None


def get_config_manager() -> ConfigManager:
    """Get global configuration manager instance."""
    global _config_manager
    if _config_manager is None:
        _config_manager = ConfigManager()
    return _config_manager


def reload_config() -> None:
    """Reload configuration from file."""
    global _config_manager
    if _config_manager is not None:
        _config_manager.load_config()