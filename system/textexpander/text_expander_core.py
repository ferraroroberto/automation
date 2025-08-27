"""
Text Expander Core Module

Handles configuration loading, validation, and core business logic
for text expansion functionality.
"""

import json
import logging
from pathlib import Path
from typing import Dict, Optional, Any, Tuple
from dataclasses import dataclass

# Configure module logger
logger = logging.getLogger(__name__)

@dataclass
class TextExpanderSettings:
    """Configuration settings for the text expander."""
    trigger_key: str
    auto_start: bool
    show_notifications: bool
    max_abbreviation_length: int
    replacement_delay_ms: int

@dataclass
class TextExpanderConfig:
    """Complete configuration for the text expander."""
    abbreviations: Dict[str, str]
    settings: TextExpanderSettings

class ConfigError(Exception):
    """Raised when configuration is invalid or missing."""
    pass

class TextExpanderCore:
    """Core business logic for text expansion functionality."""

    DEFAULT_CONFIG = {
        "abbreviations": {},
        "settings": {
            "trigger_key": "/",
            "auto_start": True,
            "show_notifications": True,
            "max_abbreviation_length": 50,
            "replacement_delay_ms": 100
        }
    }

    def __init__(self, config_path: Optional[Path] = None) -> None:
        """Initialize the text expander core.

        Args:
            config_path: Path to configuration file. If None, uses default location.
        """
        self.config_path = config_path or Path(__file__).parent / "text_expander_config.json"
        self.config: Optional[TextExpanderConfig] = None
        self._load_config()

    def _load_config(self) -> None:
        """Load and validate configuration from JSON file."""
        try:
            if not self.config_path.exists():
                logger.warning(f"📁 Configuration file not found at {self.config_path}")
                logger.info("📝 Creating default configuration")
                self._create_default_config()
                return

            with open(self.config_path, 'r', encoding='utf-8') as file:
                data = json.load(file)

            # Validate configuration structure
            self._validate_config(data)

            # Parse settings
            settings_data = data.get('settings', self.DEFAULT_CONFIG['settings'])
            settings = TextExpanderSettings(
                trigger_key=settings_data['trigger_key'],
                auto_start=settings_data['auto_start'],
                show_notifications=settings_data['show_notifications'],
                max_abbreviation_length=settings_data['max_abbreviation_length'],
                replacement_delay_ms=settings_data['replacement_delay_ms']
            )

            # Parse abbreviations
            abbreviations = data.get('abbreviations', {})

            self.config = TextExpanderConfig(
                abbreviations=abbreviations,
                settings=settings
            )

            logger.info(f"✅ Configuration loaded successfully with {len(abbreviations)} abbreviations")

        except json.JSONDecodeError as e:
            error_msg = f"❌ Invalid JSON in configuration file: {e}"
            logger.error(error_msg)
            raise ConfigError(error_msg) from e
        except KeyError as e:
            error_msg = f"❌ Missing required configuration key: {e}"
            logger.error(error_msg)
            raise ConfigError(error_msg) from e
        except Exception as e:
            error_msg = f"❌ Unexpected error loading configuration: {e}"
            logger.error(error_msg)
            raise ConfigError(error_msg) from e

    def _validate_config(self, data: Dict[str, Any]) -> None:
        """Validate configuration data structure.

        Args:
            data: Configuration data to validate.

        Raises:
            ConfigError: If configuration is invalid.
        """
        if not isinstance(data, dict):
            raise ConfigError("❌ Configuration must be a JSON object")

        # Validate abbreviations
        if 'abbreviations' in data:
            if not isinstance(data['abbreviations'], dict):
                raise ConfigError("❌ 'abbreviations' must be an object")

            for key, value in data['abbreviations'].items():
                if not isinstance(key, str) or not isinstance(value, str):
                    raise ConfigError(f"❌ Abbreviation '{key}' must have string key and value")

        # Validate settings
        if 'settings' in data:
            settings = data['settings']
            if not isinstance(settings, dict):
                raise ConfigError("❌ 'settings' must be an object")

            required_settings = ['trigger_key', 'auto_start', 'show_notifications',
                               'max_abbreviation_length', 'replacement_delay_ms']

            for setting in required_settings:
                if setting not in settings:
                    raise ConfigError(f"❌ Missing required setting: {setting}")

    def _create_default_config(self) -> None:
        """Create default configuration file."""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as file:
                json.dump(self.DEFAULT_CONFIG, file, indent=2, ensure_ascii=False)

            logger.info(f"✅ Default configuration created at {self.config_path}")

            # Load the default config
            self.config = TextExpanderConfig(
                abbreviations=self.DEFAULT_CONFIG['abbreviations'],
                settings=TextExpanderSettings(**self.DEFAULT_CONFIG['settings'])
            )

        except Exception as e:
            error_msg = f"❌ Failed to create default configuration: {e}"
            logger.error(error_msg)
            raise ConfigError(error_msg) from e

    def get_expansion(self, abbreviation: str) -> Optional[str]:
        """Get the expansion text for an abbreviation.

        Args:
            abbreviation: The abbreviation to look up (without trigger key).

        Returns:
            The expansion text if found, None otherwise.
        """
        if not self.config:
            logger.warning("⚠️ Configuration not loaded")
            return None

        full_abbreviation = f"{self.config.settings.trigger_key}{abbreviation}"
        expansion = self.config.abbreviations.get(full_abbreviation)

        if expansion:
            logger.debug(f"🔍 Found expansion for '{full_abbreviation}': {len(expansion)} characters")
        else:
            logger.debug(f"🔍 No expansion found for '{full_abbreviation}'")

        return expansion

    def add_abbreviation(self, abbreviation: str, expansion: str) -> bool:
        """Add a new abbreviation.

        Args:
            abbreviation: The abbreviation (without trigger key).
            expansion: The text to expand to.

        Returns:
            True if added successfully, False otherwise.
        """
        if not self.config:
            logger.error("❌ Configuration not loaded")
            return False

        # Validate inputs
        if not abbreviation or not expansion:
            logger.error("❌ Abbreviation and expansion cannot be empty")
            return False

        if len(abbreviation) > self.config.settings.max_abbreviation_length:
            logger.error(f"❌ Abbreviation too long (max {self.config.settings.max_abbreviation_length} chars)")
            return False

        full_abbreviation = f"{self.config.settings.trigger_key}{abbreviation}"

        # Check if abbreviation already exists
        if full_abbreviation in self.config.abbreviations:
            logger.warning(f"⚠️ Abbreviation '{full_abbreviation}' already exists")
            return False

        # Add abbreviation
        self.config.abbreviations[full_abbreviation] = expansion

        # Save configuration
        return self._save_config()

    def update_abbreviation(self, abbreviation: str, expansion: str) -> bool:
        """Update an existing abbreviation.

        Args:
            abbreviation: The abbreviation (with trigger key).
            expansion: The new expansion text.

        Returns:
            True if updated successfully, False otherwise.
        """
        if not self.config:
            logger.error("❌ Configuration not loaded")
            return False

        if abbreviation not in self.config.abbreviations:
            logger.error(f"❌ Abbreviation '{abbreviation}' not found")
            return False

        self.config.abbreviations[abbreviation] = expansion
        return self._save_config()

    def delete_abbreviation(self, abbreviation: str) -> bool:
        """Delete an abbreviation.

        Args:
            abbreviation: The abbreviation to delete (with trigger key).

        Returns:
            True if deleted successfully, False otherwise.
        """
        if not self.config:
            logger.error("❌ Configuration not loaded")
            return False

        if abbreviation not in self.config.abbreviations:
            logger.warning(f"⚠️ Abbreviation '{abbreviation}' not found")
            return False

        del self.config.abbreviations[abbreviation]
        return self._save_config()

    def _save_config(self) -> bool:
        """Save current configuration to file.

        Returns:
            True if saved successfully, False otherwise.
        """
        if not self.config:
            logger.error("❌ No configuration to save")
            return False

        try:
            # Convert config to dictionary
            config_data = {
                'abbreviations': self.config.abbreviations,
                'settings': {
                    'trigger_key': self.config.settings.trigger_key,
                    'auto_start': self.config.settings.auto_start,
                    'show_notifications': self.config.settings.show_notifications,
                    'max_abbreviation_length': self.config.settings.max_abbreviation_length,
                    'replacement_delay_ms': self.config.settings.replacement_delay_ms
                }
            }

            # Write to file
            with open(self.config_path, 'w', encoding='utf-8') as file:
                json.dump(config_data, file, indent=2, ensure_ascii=False)

            logger.info("💾 Configuration saved successfully")
            return True

        except Exception as e:
            logger.error(f"❌ Failed to save configuration: {e}")
            return False

    def get_all_abbreviations(self) -> Dict[str, str]:
        """Get all abbreviations.

        Returns:
            Dictionary of all abbreviations and their expansions.
        """
        if not self.config:
            logger.warning("⚠️ Configuration not loaded")
            return {}

        return self.config.abbreviations.copy()

    def get_settings(self) -> Optional[TextExpanderSettings]:
        """Get current settings.

        Returns:
            Current settings if configuration is loaded, None otherwise.
        """
        return self.config.settings if self.config else None
