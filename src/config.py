"""
Configuration management for StimuPop.

Provides centralized configuration with:
- YAML file loading
- Environment variable overrides
- Validation
- Sensible defaults
"""

import os
from pathlib import Path
from typing import Any, Dict, List, Optional
from dataclasses import dataclass, field

import yaml

from .exceptions import ConfigurationError


# Default configuration values
DEFAULTS: Dict[str, Any] = {
    "app": {
        "name": "StimuPop",
        "version": "2.1.0",
        "max_upload_size_mb": 200,
    },
    "images": {
        "max_size_mb": 10,
        "allowed_formats": [".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"],
        "cache_ttl_seconds": 3600,
    },
    "presentation": {
        "default_orientation": "portrait",
        "portrait_width_inches": 7.5,
        "portrait_height_inches": 10.0,
        "landscape_width_inches": 10.0,
        "landscape_height_inches": 7.5,
        "default_font_size": 14,
        "default_img_width": 5.5,
        "default_img_height": 4.0,
        "default_img_size_mode": "fit_box",
        "default_img_top": 0.5,
        "default_text_top": 5.0,
    },
    "logging": {
        "level": "INFO",
        "file": "logs/app.log",
        "max_bytes": 10485760,
        "backup_count": 5,
        "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    },
}


@dataclass
class ImageConfig:
    """Image handling configuration."""
    max_size_mb: int = 10
    allowed_formats: List[str] = field(default_factory=lambda: [".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"])
    cache_ttl_seconds: int = 3600

    @property
    def max_size_bytes(self) -> int:
        """Maximum image size in bytes."""
        return self.max_size_mb * 1024 * 1024


@dataclass
class PresentationConfig:
    """Presentation generation configuration."""
    default_orientation: str = "portrait"
    portrait_width_inches: float = 7.5
    portrait_height_inches: float = 10.0
    landscape_width_inches: float = 10.0
    landscape_height_inches: float = 7.5
    default_font_size: int = 14
    default_img_width: float = 5.5
    default_img_height: float = 4.0
    default_img_size_mode: str = "fit_box"
    default_img_top: float = 0.5
    default_text_top: float = 5.0


@dataclass
class LoggingConfig:
    """Logging configuration."""
    level: str = "INFO"
    file: str = "logs/app.log"
    max_bytes: int = 10485760
    backup_count: int = 5
    format: str = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"


@dataclass
class AppConfig:
    """Application metadata configuration."""
    name: str = "StimuPop"
    version: str = "2.1.0"
    max_upload_size_mb: int = 200

    @property
    def max_upload_size_bytes(self) -> int:
        """Maximum upload size in bytes."""
        return self.max_upload_size_mb * 1024 * 1024


class Config:
    """
    Central configuration manager.

    Loads configuration from:
    1. Default values
    2. YAML config file (if exists)
    3. Environment variables (override)

    Environment variable format: APP_SECTION_KEY
    Example: APP_IMAGES_MAX_SIZE_MB=20

    Usage:
        config = Config()  # Uses default config path
        config = Config("path/to/config.yaml")

        # Access settings
        print(config.images.max_size_mb)
        print(config.security.block_private_ips)
    """

    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize configuration.

        Args:
            config_path: Path to YAML config file. If None, looks for
                        'config.yaml' in the project root.
        """
        self._raw_config = self._load_config(config_path)
        self._apply_env_overrides()
        self._validate()

        # Create typed config objects
        self.app = AppConfig(**self._raw_config.get("app", {}))
        self.images = ImageConfig(**self._raw_config.get("images", {}))
        self.presentation = PresentationConfig(**self._raw_config.get("presentation", {}))
        self.logging = LoggingConfig(**self._raw_config.get("logging", {}))

    def _load_config(self, config_path: Optional[str]) -> Dict[str, Any]:
        """Load configuration from YAML file with defaults."""
        config = self._deep_copy(DEFAULTS)

        if config_path is None:
            # Look for config.yaml in project root
            project_root = Path(__file__).parent.parent
            config_path = project_root / "config.yaml"
        else:
            config_path = Path(config_path)

        if config_path.exists():
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    yaml_config = yaml.safe_load(f) or {}
                config = self._deep_merge(config, yaml_config)
            except yaml.YAMLError as e:
                raise ConfigurationError(
                    f"Invalid YAML syntax in config file",
                    setting=str(config_path),
                    details=str(e)
                )
            except IOError as e:
                raise ConfigurationError(
                    f"Cannot read config file",
                    setting=str(config_path),
                    details=str(e)
                )

        return config

    def _apply_env_overrides(self) -> None:
        """Apply environment variable overrides."""
        for section, settings in self._raw_config.items():
            if not isinstance(settings, dict):
                continue
            for key, value in settings.items():
                env_key = f"APP_{section.upper()}_{key.upper()}"
                env_value = os.environ.get(env_key)
                if env_value is not None:
                    # Convert to appropriate type
                    self._raw_config[section][key] = self._convert_type(
                        env_value, type(value)
                    )

    def _convert_type(self, value: str, target_type: type) -> Any:
        """Convert string value to target type."""
        if target_type == bool:
            return value.lower() in ("true", "1", "yes", "on")
        elif target_type == int:
            return int(value)
        elif target_type == float:
            return float(value)
        elif target_type == list:
            # Comma-separated values
            return [v.strip() for v in value.split(",")]
        return value

    def _validate(self) -> None:
        """Validate configuration values."""
        # Validate image settings
        images = self._raw_config.get("images", {})
        if images.get("max_size_mb", 0) <= 0:
            raise ConfigurationError(
                "Must be greater than 0",
                setting="images.max_size_mb"
            )

        # Validate logging level
        logging_config = self._raw_config.get("logging", {})
        valid_levels = ["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]
        if logging_config.get("level", "INFO").upper() not in valid_levels:
            raise ConfigurationError(
                f"Must be one of: {', '.join(valid_levels)}",
                setting="logging.level"
            )

    @staticmethod
    def _deep_copy(d: Dict[str, Any]) -> Dict[str, Any]:
        """Create a deep copy of a dictionary."""
        result = {}
        for key, value in d.items():
            if isinstance(value, dict):
                result[key] = Config._deep_copy(value)
            elif isinstance(value, list):
                result[key] = value.copy()
            else:
                result[key] = value
        return result

    @staticmethod
    def _deep_merge(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
        """Deep merge two dictionaries, with override taking precedence."""
        result = Config._deep_copy(base)
        for key, value in override.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = Config._deep_merge(result[key], value)
            else:
                result[key] = value
        return result


# Global config instance (lazy loaded)
_config: Optional[Config] = None


def get_config(config_path: Optional[str] = None) -> Config:
    """
    Get the global configuration instance.

    Creates the instance on first call. Subsequent calls return
    the same instance unless a different config_path is provided.

    Args:
        config_path: Optional path to config file. Only used on first call
                    or if explicitly providing a new path.

    Returns:
        Config instance
    """
    global _config
    if _config is None or config_path is not None:
        _config = Config(config_path)
    return _config
