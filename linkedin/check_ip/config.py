#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Configuration Management Module

This module handles configuration loading from JSON files and command-line argument parsing
for the LinkedIn image search application.

Author: Roberto (Refactored by Claude)
Date: March 2025
"""

import argparse
import json
import logging
import os
import sys
from pathlib import Path
from typing import Dict, Any, Optional
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

logger = logging.getLogger(__name__)


class Config:
    """Configuration manager for the LinkedIn image search application."""
    
    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize configuration manager.
        
        Args:
            config_path (Optional[str]): Path to configuration file
        """
        self.config_path = self._resolve_config_path(config_path)
        self.config_data = self._load_config()
        
    def _resolve_config_path(self, config_path: Optional[str]) -> Path:
        """
        Resolve the configuration file path.
        
        Args:
            config_path (Optional[str]): User-provided config path
            
        Returns:
            Path: Resolved configuration file path
        """
        if config_path is None:
            # Default to linkedin_search_image.json in the same directory as this script
            script_dir = Path(__file__).parent
            return script_dir / "linkedin_search_image.json"
        return Path(config_path)
    
    def _create_default_config(self) -> Dict[str, Any]:
        """
        Create default configuration structure.
        
        Returns:
            Dict[str, Any]: Default configuration dictionary
        """
        return {
            "api_keys": {
                "imgur_client_id": "",
                "imgur_access_token": "",
                "serpapi_key": ""
            },
            "folders": {
                "images_folder": "./images",
                "metadata_folder": "./metadata"
            },
            "file_names": {
                "metadata_file": "metadata.xlsx",
                "results_file": "search_results_database.xlsx",
                "api_history_file": "api_history_database.xlsx"
            },
            "processing_thresholds": {
                "linkedin_high_threshold": 50,
                "high_linkedin_days": 7,
                "standard_days": 180
            },
            "search_settings": {
                "do_exact_search": True,
                "do_similar_search": False
            },
            "allowed_extensions": [".jpg", ".jpeg", ".png", ".gif", ".bmp"],
            "api_endpoints": {
                "search_endpoint": "https://serpapi.com/search?engine=google_lens"
            }
        }
    
    def _load_config(self) -> Dict[str, Any]:
        """
        Load configuration from JSON file.
        
        Returns:
            Dict[str, Any]: Configuration dictionary
            
        Raises:
            SystemExit: If configuration file is missing or invalid
        """
        if not self.config_path.exists():
            logger.error(f"❌ Configuration file not found: {self.config_path}")
            logger.info("Creating a default configuration file...")
            
            default_config = self._create_default_config()
            
            try:
                with open(self.config_path, 'w') as f:
                    json.dump(default_config, f, indent=2)
                logger.error("Please update the configuration file with your API keys and paths.")
                sys.exit(1)
            except Exception as e:
                logger.error(f"❌ Error creating default configuration file: {e}")
                sys.exit(1)
        
        try:
            with open(self.config_path, 'r') as f:
                config = json.load(f)
            logger.info(f"✅ Configuration loaded from: {self.config_path}")
            
            # Resolve environment variables in the configuration
            config = self._resolve_env_vars(config)
            
            return config
        except json.JSONDecodeError as e:
            logger.error(f"❌ Error parsing configuration file: {e}")
            sys.exit(1)
        except Exception as e:
            logger.error(f"❌ Error loading configuration file: {e}")
            sys.exit(1)
    
    def _resolve_env_vars(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Resolve environment variables in configuration values.
        
        Args:
            config (Dict[str, Any]): Configuration dictionary
            
        Returns:
            Dict[str, Any]: Configuration with resolved environment variables
        """
        def resolve_value(value):
            if isinstance(value, str) and value.startswith("${") and value.endswith("}"):
                env_var = value[2:-1]  # Remove ${ and }
                env_value = os.getenv(env_var)
                if env_value is None:
                    logger.warning(f"⚠️ Environment variable {env_var} not found, using empty string")
                    return ""
                return env_value
            elif isinstance(value, dict):
                return {k: resolve_value(v) for k, v in value.items()}
            elif isinstance(value, list):
                return [resolve_value(item) for item in value]
            else:
                return value
        
        return resolve_value(config)
    
    def get_api_keys(self) -> Dict[str, str]:
        """Get API keys from configuration."""
        return self.config_data["api_keys"]
    
    def get_folders(self) -> Dict[str, str]:
        """Get folder paths from configuration."""
        return self.config_data["folders"]
    
    def get_file_names(self) -> Dict[str, str]:
        """Get file names from configuration."""
        return self.config_data["file_names"]
    
    def get_processing_thresholds(self) -> Dict[str, int]:
        """Get processing thresholds from configuration."""
        return self.config_data["processing_thresholds"]
    
    def get_search_settings(self) -> Dict[str, bool]:
        """Get search settings from configuration."""
        return self.config_data["search_settings"]
    
    def get_allowed_extensions(self) -> list:
        """Get allowed file extensions from configuration."""
        return self.config_data["allowed_extensions"]
    
    def get_api_endpoints(self) -> Dict[str, str]:
        """Get API endpoints from configuration."""
        return self.config_data["api_endpoints"]


def parse_arguments() -> argparse.Namespace:
    """
    Parse command line arguments with sensible defaults from configuration.
    
    Returns:
        argparse.Namespace: Parsed command line arguments
    """
    # Load configuration first to use as defaults
    config = Config()
    
    parser = argparse.ArgumentParser(
        description="Enhanced Image Search Script with metadata tracking and API history",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    # Configuration file
    parser.add_argument(
        "--config",
        type=str,
        help="Path to configuration JSON file (default: linkedin_search_image.json in script directory)"
    )
    
    # Folder paths
    parser.add_argument(
        "--images-folder",
        type=str,
        default=config.get_folders()["images_folder"],
        help="Folder containing the images to process"
    )
    parser.add_argument(
        "--metadata-folder",
        type=str,
        default=config.get_folders()["metadata_folder"],
        help="Folder for storing Excel files and metadata"
    )
    
    # API Keys
    api_keys = config.get_api_keys()
    parser.add_argument(
        "--imgur-client-id",
        type=str,
        default=api_keys["imgur_client_id"],
        help="Imgur Client ID"
    )
    parser.add_argument(
        "--imgur-access-token",
        type=str,
        default=api_keys["imgur_access_token"],
        help="Imgur API access token"
    )
    parser.add_argument(
        "--serpapi-key",
        type=str,
        default=api_keys["serpapi_key"],
        help="SerpApi API key"
    )
    
    # File names
    file_names = config.get_file_names()
    parser.add_argument(
        "--metadata-file",
        type=str,
        default=file_names["metadata_file"],
        help="Name of the metadata Excel file"
    )
    parser.add_argument(
        "--results-file",
        type=str,
        default=file_names["results_file"],
        help="Name of the results database Excel file"
    )
    parser.add_argument(
        "--api-history-file",
        type=str,
        default=file_names["api_history_file"],
        help="Name of the API history Excel file"
    )
    
    # Processing thresholds
    thresholds = config.get_processing_thresholds()
    parser.add_argument(
        "--linkedin-high-threshold",
        type=int,
        default=thresholds["linkedin_high_threshold"],
        help="LinkedIn count threshold for high-frequency processing"
    )
    parser.add_argument(
        "--high-linkedin-days",
        type=int,
        default=thresholds["high_linkedin_days"],
        help="Days between processing for images with high LinkedIn count"
    )
    parser.add_argument(
        "--standard-days",
        type=int,
        default=thresholds["standard_days"],
        help="Days between processing for standard images"
    )
    
    # Search controls
    search_settings = config.get_search_settings()
    parser.add_argument(
        "--skip-exact-search",
        action="store_true",
        default=not search_settings["do_exact_search"],
        help="Skip exact matches search"
    )
    parser.add_argument(
        "--enable-similar-search",
        action="store_true",
        default=search_settings["do_similar_search"],
        help="Enable similar matches search"
    )
    
    args = parser.parse_args()
    
    # If a different config file was specified, reload configuration
    if args.config:
        config = Config(args.config)
    
    # Convert folder paths to Path objects
    args.images_folder = Path(args.images_folder)
    args.metadata_folder = Path(args.metadata_folder)
    
    # Store the config object in args for later use
    args.config_object = config
    
    return args


def get_app_config(args: argparse.Namespace) -> Dict[str, Any]:
    """
    Build application configuration from parsed arguments and config file.
    
    Args:
        args (argparse.Namespace): Parsed command line arguments
        
    Returns:
        Dict[str, Any]: Complete application configuration
    """
    config = args.config_object
    
    return {
        # API Keys
        "imgur_client_id": args.imgur_client_id,
        "imgur_access_token": args.imgur_access_token,
        "serpapi_key": args.serpapi_key,
        
        # Paths
        "images_folder": args.images_folder,
        "metadata_folder": args.metadata_folder,
        "metadata_file": args.metadata_file,
        "results_file": args.results_file,
        "api_history_file": args.api_history_file,
        
        # Processing settings
        "linkedin_high_threshold": args.linkedin_high_threshold,
        "high_linkedin_days": args.high_linkedin_days,
        "standard_days": args.standard_days,
        
        # Search settings
        "do_exact_search": not args.skip_exact_search,
        "do_similar_search": args.enable_similar_search,
        
        # From config file
        "allowed_extensions": set(config.get_allowed_extensions()),
        "search_endpoint": config.get_api_endpoints()["search_endpoint"]
    } 