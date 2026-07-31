"""
Configuration and logging for the photo-archive tool.

Owns `photos_archive.json`: loading it, validating it, filling in defaults, and
setting up the run's log file. Split out of `photos_archive.py` (audit issue
#92) so the metadata, duplicate-detection and CLI layers share one config seam
instead of a module-level global.

The config is loaded lazily on the first `get_config()` call rather than at
import time, so a malformed or missing `photos_archive.json` surfaces inside
the caller's error handling instead of as a bare import-time traceback.
"""

import json
import logging
import os
from datetime import datetime
from typing import Any, Dict, Optional

CONFIG_FILENAME = 'photos_archive.json'

_config: Optional[Dict[str, Any]] = None


def load_config() -> Dict[str, Any]:
    """Load and validate configuration from photos_archive.json."""
    config_path = os.path.join(os.path.dirname(__file__), CONFIG_FILENAME)
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)

        # Validate required fields
        required_fields = ['source_folder', 'destination_folder']
        for field in required_fields:
            if field not in config:
                raise ValueError(f"Required field missing: {field}")

        # Warn if configured paths don't exist
        for folder_field in ['source_folder', 'destination_folder']:
            folder_path = config.get(folder_field)
            if folder_path and not os.path.exists(folder_path):
                logging.warning(f"Path not found: {folder_field} = {folder_path}")

        # Validate data types
        thresholds = config.get('processing_thresholds', {})
        if 'short_sequence_seconds' in thresholds and not isinstance(thresholds['short_sequence_seconds'], int):
            raise ValueError("short_sequence_seconds must be integer")

        # Set defaults for optional fields
        config.setdefault('log_folder', config.get('destination_folder'))
        config['logging'] = config.get('logging', {})
        config['logging'].setdefault('level', 'INFO')
        config['logging'].setdefault('progress_reporting_interval', 1000)
        config['duplicate_detection'] = config.get('duplicate_detection', {})
        config['duplicate_detection'].setdefault('enable_sha256_check', True)
        config['duplicate_detection'].setdefault('keep_largest_file', True)

        return config
    except FileNotFoundError:
        raise FileNotFoundError(f"Config file not found: {config_path}")
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON: {e}")
    except Exception as e:
        raise ValueError(f"Config load error: {e}")


def get_config() -> Dict[str, Any]:
    """Return the validated config, loading it on first use."""
    global _config
    if _config is None:
        _config = load_config()
    return _config


def setup_logging(log_folder: Optional[str] = None) -> None:
    """Configure logging to timestamped file in specified folder."""
    config = get_config()
    if log_folder is None:
        log_folder = config.get('log_folder', config['destination_folder'])

    timestamp = datetime.now().strftime('%Y%m%d-%H%M')
    log_file = os.path.join(log_folder, f'photo_processing_{timestamp}.log')

    os.makedirs(log_folder, exist_ok=True)

    # Set logging level from config
    log_level = getattr(logging, config['logging']['level'].upper(), logging.INFO)

    fmt = '%(asctime)s - %(levelname)s - %(message)s'
    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)
    file_handler = logging.FileHandler(log_file, mode='w')
    file_handler.setFormatter(logging.Formatter(fmt))
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter(fmt))
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)
    logging.info("Logging initialized: %s", log_file)
