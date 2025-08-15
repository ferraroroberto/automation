#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Enhanced Image Search Script with Metadata, Cumulative Database, and API History

This script performs image searches using Google Lens via the SerpAPI.
It includes the following enhancements:
- Metadata management: Tracks images already processed and their Imgur URLs
- Cumulative database: Preserves historical search results and user annotations
- API history database: Stores complete details of all API calls including raw JSON responses
- Optimization: Avoids re-uploading images that were previously processed
- Social media extraction: Extracts and formats post dates from various platforms
- Time-based processing: Skips images processed within the last 10 days
- Configurable search types: Control whether to do exact matches or enable similar matches

Author: Roberto (Enhanced by Claude)
Date: March 2025
"""

import argparse
import datetime
from datetime import timedelta, timezone
import json
import logging
import numpy as np
import os
import pandas as pd
from pathlib import Path
import re
import requests
from typing import Dict, List, Optional, Tuple, Union

import sys
import time

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def load_config(config_path: Optional[str] = None) -> dict:
    """
    Load configuration from JSON file.
    
    Args:
        config_path (Optional[str]): Path to the configuration file
        
    Returns:
        dict: Configuration dictionary
    """
    if config_path is None:
        # Default to linkedin_search_image.json in the same directory as this script
        script_dir = Path(__file__).parent
        config_path = script_dir / "linkedin_search_image.json"
    else:
        config_path = Path(config_path)
    
    if not config_path.exists():
        logger.error(f"❌ Configuration file not found: {config_path}")
        logger.info("Creating a default configuration file...")
        
        # Create default configuration
        default_config = {
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
        
        # Save default configuration
        with open(config_path, 'w') as f:
            json.dump(default_config, f, indent=2)
        
        logger.error("Please update the configuration file with your API keys and paths.")
        sys.exit(1)
    
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
        logger.info(f"✅ Configuration loaded from: {config_path}")
        return config
    except json.JSONDecodeError as e:
        logger.error(f"❌ Error parsing configuration file: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"❌ Error loading configuration file: {e}")
        sys.exit(1)

def parse_arguments():
    """
    Parse command line arguments with sensible defaults from configuration.
    """
    # Load configuration first to use as defaults
    config = load_config()
    
    parser = argparse.ArgumentParser(
        description="Enhanced Image Search Script with metadata tracking and API history",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    # Configuration file
    parser.add_argument(
        "--config",
        type=str,
        help="Path to configuration JSON file (default: config.json in script directory)"
    )
    
    # Folder paths
    parser.add_argument(
        "--images-folder",
        type=str,
        default=config["folders"]["images_folder"],
        help="Folder containing the images to process"
    )
    parser.add_argument(
        "--metadata-folder",
        type=str,
        default=config["folders"]["metadata_folder"],
        help="Folder for storing Excel files and metadata"
    )
    
    # API Keys
    parser.add_argument(
        "--imgur-client-id",
        type=str,
        default=config["api_keys"]["imgur_client_id"],
        help="Imgur Client ID"
    )
    parser.add_argument(
        "--imgur-access-token",
        type=str,
        default=config["api_keys"]["imgur_access_token"],
        help="Imgur API access token"
    )
    parser.add_argument(
        "--serpapi-key",
        type=str,
        default=config["api_keys"]["serpapi_key"],
        help="SerpApi API key"
    )
    
    # File names
    parser.add_argument(
        "--metadata-file",
        type=str,
        default=config["file_names"]["metadata_file"],
        help="Name of the metadata Excel file"
    )
    parser.add_argument(
        "--results-file",
        type=str,
        default=config["file_names"]["results_file"],
        help="Name of the results database Excel file"
    )
    parser.add_argument(
        "--api-history-file",
        type=str,
        default=config["file_names"]["api_history_file"],
        help="Name of the API history Excel file"
    )
    
    # Processing thresholds
    parser.add_argument(
        "--linkedin-high-threshold",
        type=int,
        default=config["processing_thresholds"]["linkedin_high_threshold"],
        help="LinkedIn count threshold for high-frequency processing"
    )
    parser.add_argument(
        "--high-linkedin-days",
        type=int,
        default=config["processing_thresholds"]["high_linkedin_days"],
        help="Days between processing for images with high LinkedIn count"
    )
    parser.add_argument(
        "--standard-days",
        type=int,
        default=config["processing_thresholds"]["standard_days"],
        help="Days between processing for standard images"
    )
    
    # Search controls
    parser.add_argument(
        "--skip-exact-search",
        action="store_true",
        default=not config["search_settings"]["do_exact_search"],
        help="Skip exact matches search"
    )
    parser.add_argument(
        "--enable-similar-search",
        action="store_true",
        default=config["search_settings"]["do_similar_search"],
        help="Enable similar matches search"
    )
    
    args = parser.parse_args()
    
    # If a different config file was specified, reload configuration
    if args.config:
        config = load_config(args.config)
        # Update defaults with new config (but command line args take precedence)
        # This is already handled by argparse since we parse after loading config
    
    # Convert folder paths to Path objects
    args.images_folder = Path(args.images_folder)
    args.metadata_folder = Path(args.metadata_folder)
    
    # Store the loaded config in args for later use
    args.config_data = config
    
    return args

# Global variables will be set based on command line arguments
def initialize_globals(args):
    """
    Initialize global variables based on command line arguments.
    """
    global IMGUR_CLIENT_ID, ACCESS_TOKEN, SERPAPI_KEY
    global SEARCH_ENDPOINT
    global METADATA_FOLDER, IMAGES_FOLDER
    global METADATA_FILE, RESULTS_FILE, API_HISTORY_FILE
    global LINKEDIN_HIGH_THRESHOLD, HIGH_LINKEDIN_THRESHOLD_DAYS, STANDARD_THRESHOLD_DAYS
    global DO_EXACT_SEARCH, DO_SIMILAR_SEARCH
    global ALLOWED_EXT
    
    # API Keys
    IMGUR_CLIENT_ID = args.imgur_client_id
    ACCESS_TOKEN = args.imgur_access_token
    SERPAPI_KEY = args.serpapi_key
    
    # API endpoint from config
    SEARCH_ENDPOINT = args.config_data["api_endpoints"]["search_endpoint"]
    
    # Folders
    METADATA_FOLDER = args.metadata_folder
    IMAGES_FOLDER = args.images_folder
    
    # File names
    METADATA_FILE = args.metadata_file
    RESULTS_FILE = args.results_file
    API_HISTORY_FILE = args.api_history_file
    
    # Thresholds
    LINKEDIN_HIGH_THRESHOLD = args.linkedin_high_threshold
    HIGH_LINKEDIN_THRESHOLD_DAYS = args.high_linkedin_days
    STANDARD_THRESHOLD_DAYS = args.standard_days
    
    # Search controls
    DO_EXACT_SEARCH = not args.skip_exact_search
    DO_SIMILAR_SEARCH = args.enable_similar_search
    
    # Allowed extensions from config
    ALLOWED_EXT = set(args.config_data["allowed_extensions"])

def setup_folders_and_files() -> Tuple[Path, Path, Path, Path]:
    """
    Setup folders and file paths.
    
    Returns:
        Tuple[Path, Path, Path, Path]: Paths for images folder, metadata file, results file, and API history file
    """
    # Ensure both folders exist
    if not IMAGES_FOLDER.exists():
        logger.error(f"❌  Error: Images folder '{IMAGES_FOLDER}' does not exist.")
        raise FileNotFoundError(f"Images folder not found: {IMAGES_FOLDER}")
    
    # Create metadata folder if it doesn't exist
    METADATA_FOLDER.mkdir(parents=True, exist_ok=True)
    logger.info(f"✅ Images folder exists: {IMAGES_FOLDER}")
    logger.info(f"✅ Metadata folder exists/created: {METADATA_FOLDER}")
    
    # Define paths for metadata, results, and API history files in the metadata folder
    metadata_path = METADATA_FOLDER / METADATA_FILE
    results_path = METADATA_FOLDER / RESULTS_FILE
    api_history_path = METADATA_FOLDER / API_HISTORY_FILE
    
    return IMAGES_FOLDER, metadata_path, results_path, api_history_path

def load_metadata(metadata_path: Path) -> pd.DataFrame:
    """
    Load metadata file if it exists, or create an empty DataFrame.
    
    Args:
        metadata_path (Path): Path to the metadata file
        
    Returns:
        pd.DataFrame: Metadata DataFrame with images and their Imgur URLs
    """
    # Define the essential columns we require
    essential_columns = ['filename', 'last_processed_date', 'imgur_url']
    
    # Define count columns we'll add/update
    count_columns = [
        'total_links', 'exact_match_count', 'similar_match_count',
        'linkedin_count', 'instagram_count', 'twitter_count', 
        'facebook_count', 'pinterest_count'
    ]
    
    if metadata_path.exists():
        try:
            logger.info(f"📊 Loading existing metadata from {metadata_path}")
            metadata_df = pd.read_excel(metadata_path)
            
            # Validate essential columns
            if not all(col in metadata_df.columns for col in essential_columns):
                logger.warning("⚠️ Metadata file is missing essential columns. Creating backup and initializing new file.")
                # Create backup of existing file
                backup_path = metadata_path.with_suffix(f".bak.{int(time.time())}.xlsx")
                metadata_df.to_excel(backup_path, index=False)
                # Create new DataFrame with essential columns
                metadata_df = pd.DataFrame(columns=essential_columns)
            else:
                # Add any missing count columns (for backward compatibility)
                for col in count_columns:
                    if col not in metadata_df.columns:
                        # Insert count columns after the essential columns
                        insert_pos = len(essential_columns)
                        metadata_df.insert(insert_pos, col, 0)
                        logger.info(f"📊 Added missing column '{col}' to metadata")
            
            return metadata_df
        except Exception as e:
            logger.error(f"❌ Error loading metadata file: {e}")
            logger.info("🔄 Initializing new metadata file")
            # Start with essential columns
            df = pd.DataFrame(columns=essential_columns)
            # Add count columns
            for col in count_columns:
                df[col] = 0
            return df
    else:
        logger.info("📊 No existing metadata found. Initializing new metadata file.")
        # Start with essential columns
        df = pd.DataFrame(columns=essential_columns)
        # Add count columns
        for col in count_columns:
            df[col] = 0
        return df

def load_results_database(results_path: Path) -> pd.DataFrame:
    """
    Load existing results database if it exists, or create an empty DataFrame.
    
    Args:
        results_path (Path): Path to the results database file
        
    Returns:
        pd.DataFrame: Results database with all previous search results
    """
    # Define the required columns for the results database
    required_columns = ['local_image', 'found_link']
    
    # Define standard columns that should be present
    standard_columns = [
        'local_image', 'uploaded_url', 'found_link', 
        'title', 'duplicate', 'match_type', 'source', 
        'post_date', 'search_date', 'order'
    ]
    
    if results_path.exists():
        try:
            logger.info(f"📊 Loading existing results database from {results_path}")
            results_df = pd.read_excel(results_path)
            
            # Validate that we have at least the required columns
            if not all(col in results_df.columns for col in required_columns):
                logger.warning("⚠️ Results file is missing required columns. Creating backup and initializing new file.")
                # Create backup of existing file
                backup_path = results_path.with_suffix(f".bak.{int(time.time())}.xlsx")
                results_df.to_excel(backup_path, index=False)
                # Create new DataFrame with standard columns
                results_df = pd.DataFrame(columns=standard_columns)
            
            # Handle the transition from 'snippet' to 'duplicate' column
            if 'snippet' in results_df.columns and 'duplicate' not in results_df.columns:
                logger.info("🔄 Converting 'snippet' column to 'duplicate' column")
                # Rename the column
                results_df = results_df.rename(columns={'snippet': 'duplicate'})
                # Initialize all values to 0
                results_df['duplicate'] = 0
            elif 'snippet' in results_df.columns and 'duplicate' in results_df.columns:
                # Both columns exist, drop snippet
                logger.info("🔄 Dropping 'snippet' column as 'duplicate' already exists")
                results_df = results_df.drop(columns=['snippet'])
            elif 'duplicate' not in results_df.columns:
                # Add duplicate column if it doesn't exist
                logger.info("🔄 Adding 'duplicate' column")
                results_df['duplicate'] = 0
            
            return results_df
        except Exception as e:
            logger.error(f"❌ Error loading results file: {e}")
            logger.info("🔄 Initializing new results database")
            return pd.DataFrame(columns=standard_columns)
    else:
        logger.info("📊 No existing results database found. Initializing new database.")
        return pd.DataFrame(columns=standard_columns)

def load_api_history_database(history_path: Path) -> pd.DataFrame:
    """
    Load existing API history database if it exists, or create an empty DataFrame.
    
    Args:
        history_path (Path): Path to the API history database file
        
    Returns:
        pd.DataFrame: API history database with all previous API calls
    """
    # Define the required columns for API history
    required_columns = ['search_type', 'local_image', 'imgur_url', 'search_date']
    
    # Define standard columns for API history
    standard_columns = [
        'search_type', 'local_image', 'imgur_url', 'search_date',
        'api_id', 'api_status', 'json_endpoint', 'created_at', 'processed_at',
        'total_time_taken', 'engine', 'url', 'page_token', 'raw_json',
        'search_engine_query', 'raw_html_file'
    ]
    
    # Optional columns that might be present depending on API response
    optional_columns = [
        'user', 'requester_ip', 'rtt', 'device', 'source_user_agent', 
        'used_user_agent', 'proxy_provider', 'playground_link'
    ]
    
    if history_path.exists():
        try:
            logger.info(f"📊 Loading existing API history database from {history_path}")
            history_df = pd.read_excel(history_path)
            
            # Validate that we have at least required columns
            if not all(col in history_df.columns for col in required_columns):
                logger.warning("⚠️ API history file is missing required columns. Creating backup and initializing new file.")
                # Create backup of existing file
                backup_path = history_path.with_suffix(f".bak.{int(time.time())}.xlsx")
                history_df.to_excel(backup_path, index=False)
                # Create new DataFrame with all columns
                all_columns = standard_columns + optional_columns
                history_df = pd.DataFrame(columns=all_columns)
            
            return history_df
        except Exception as e:
            logger.error(f"❌ Error loading API history file: {e}")
            logger.info("🔄 Initializing new API history database")
            all_columns = standard_columns + optional_columns
            return pd.DataFrame(columns=all_columns)
    else:
        logger.info("📊 No existing API history database found. Initializing new database.")
        all_columns = standard_columns + optional_columns
        return pd.DataFrame(columns=all_columns)

# Function to upload image to Imgur (authenticated)
def upload_to_imgur(image_path: Path, retries=20):
    """
    Uploads an image to Imgur and retries if a 503/429 error occurs using exponential backoff.
    
    Args:
        image_path (Path): Path to the image file
        retries (int): Number of retry attempts
        
    Returns:
        Optional[str]: Imgur URL if successful, None otherwise
    """
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}
    delay = 5  # Initial retry delay in seconds

    for attempt in range(1, retries + 1):  # Start from attempt 1
        try:
            logger.info(f"🚀 Attempt {attempt}/{retries}: Uploading {image_path.name}")

            with open(image_path, "rb") as file:
                response = requests.post(
                    "https://api.imgur.com/3/image",
                    headers=headers,
                    files={"image": file}
                )

            if response.status_code == 200:
                img_url = response.json()["data"]["link"]
                logger.info(f"✅ Image uploaded successfully: {img_url}")
                return img_url
            
            elif response.status_code == 503:
                logger.warning(f"⚠️ Server busy (503 Error). Attempt {attempt}/{retries}. Retrying in {delay} seconds...")
            
            elif response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", delay))  # Default to delay if not provided
                logger.warning(f"⚠️ Rate limit exceeded (429 Error). Attempt {attempt}/{retries}. Waiting {retry_after} seconds before retrying...")
                delay = retry_after  # Use API-provided delay instead of exponential

            else:
                logger.error(f"❌ Error {response.status_code}: {response.text}. Upload failed.")
                return None
        
        except requests.exceptions.RequestException as e:
            logger.error(f"❌ Network error: {str(e)}. Attempt {attempt}/{retries}. Retrying in {delay} seconds...")

        # Apply exponential backoff, doubling the delay each attempt
        time.sleep(delay)
        delay = min(delay * 2, 1800)  # Max delay capped at 30 minutes (to prevent infinite waiting)

    logger.error(f"❌ Upload failed after {retries} attempts. Skipping {image_path.name}.")
    return None

def add_to_api_history(
    history_df: pd.DataFrame,
    search_type: str,
    local_image: str,
    imgur_url: str,
    response_data: dict,
    page_token: Optional[str] = None,
    search_type_param: Optional[str] = None
) -> pd.DataFrame:
    """
    Add an API call record to the API history database.
    
    Args:
        history_df (pd.DataFrame): API history database
        search_type (str): Type of search (human-readable description)
        local_image (str): Filename of the local image
        imgur_url (str): Imgur URL of the uploaded image
        response_data (dict): API response data
        page_token (Optional[str]): Page token for legacy searches (deprecated)
        search_type_param (Optional[str]): The actual 'type' parameter used in the API call
        
    Returns:
        pd.DataFrame: Updated API history database
    """
    try:
        # Extract metadata from response
        search_metadata = response_data.get('search_metadata', {})
        search_parameters = response_data.get('search_parameters', {})
        
        # Create a new record
        new_record = {
            'search_type': search_type,
            'local_image': local_image,
            'imgur_url': imgur_url,
            'search_date': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            
            # API metadata
            'api_id': search_metadata.get('id'),
            'api_status': search_metadata.get('status'),
            'json_endpoint': search_metadata.get('json_endpoint'),
            'created_at': search_metadata.get('created_at'),
            'processed_at': search_metadata.get('processed_at'),
            'total_time_taken': search_metadata.get('total_time_taken'),
            'search_engine_query': search_metadata.get('google_lens_url'),
            'raw_html_file': search_metadata.get('raw_html_file'),
            
            # Search parameters
            'engine': search_parameters.get('engine'),
            'url': search_parameters.get('url'),
            'page_token': page_token,  # Keep for backward compatibility, will be None for new searches
            
            # Additional SerpAPI-specific fields that might be available (set to None if not available)
            'user': 'roberto.ferraro@gmail.com',  # User email is not in the API response, but we know it
            'requester_ip': None,  # Not available in API response
            'rtt': 1,  # Not available in API response, default to 1
            'device': 'desktop',  # Not available in API response, default to desktop
            'source_user_agent': 'python-requests/2.28.2',  # Not available in API response
            'used_user_agent': None,  # Not available in API response
            'proxy_provider': None,  # Not available in API response
            'playground_link': None,  # We'll construct this below
            
            # Store the complete response for future reference
            'raw_json': json.dumps(response_data)
        }
        
        # Add the search type parameter to the record if provided
        if search_type_param:
            new_record['search_type_param'] = search_type_param
        
        # Construct playground link based on search type
        url_param = urllib.parse.quote(imgur_url)
        if search_type_param:
            new_record['playground_link'] = f"https://serpapi.com/playground?engine=google_lens&url={url_param}&type={search_type_param}"
        else:
            new_record['playground_link'] = f"https://serpapi.com/playground?engine=google_lens&url={url_param}"
        
        # Add to history database
        return pd.concat([history_df, pd.DataFrame([new_record])], ignore_index=True)
    except Exception as e:
        logger.error(f"❌ Error adding to API history: {str(e)}")
        return history_df

def search_google_lens(image_url: str, local_image: str, api_history_df: pd.DataFrame, search_type: str = "all") -> Tuple[Optional[dict], pd.DataFrame]:
    """
    Searches Google Lens via SerpApi with the specified search type.
    
    Args:
        image_url (str): URL of the image to search
        local_image (str): Filename of the local image
        api_history_df (pd.DataFrame): API history database
        search_type (str): Type of search - "all", "exact_matches", "visual_matches", or "products"
        
    Returns:
        Tuple[Optional[dict], pd.DataFrame]: Data and updated API history
    """
    params = {
        "engine": "google_lens",
        "url": image_url,
        "type": search_type,
        "api_key": SERPAPI_KEY,
        "no_cache": True  # Forces fresh results
    }

    try:
        response = requests.get(SEARCH_ENDPOINT, params=params)

        logger.info(f"🔍 Response Status for {search_type}: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            
            # Add to API history
            api_history_df = add_to_api_history(
                api_history_df,
                f'{search_type.replace("_", " ").title()} Search',
                local_image,
                image_url,
                data,
                page_token=None,
                search_type_param=search_type
            )
            
            return data, api_history_df
        else:
            logger.error(f"❌ SerpApi Request Failed: {response.status_code}")
            return None, api_history_df
    except Exception as e:
        logger.error(f"❌ Exception during Google Lens search: {str(e)}")
        return None, api_history_df

def extract_post_date(url: str) -> Optional[str]:
    """
    Extracts post date from Twitter/X & LinkedIn URLs. Returns None if not applicable.
    
    Args:
        url (str): URL to extract date from
        
    Returns:
        Optional[str]: Formatted date string or None if extraction fails
    """
    # Twitter/X URL Format: https://x.com/FerraroRoberto/status/1598186442865491971
    twitter_match = re.search(r"(?:twitter|x)\.com/.*/status/(\d+)", url)
    
    if twitter_match:
        try:
            tweet_id = int(twitter_match.group(1))
            tweet_timestamp = (tweet_id >> 22) + 1288834974657  # Twitter Snowflake Epoch
            return datetime.datetime.fromtimestamp(tweet_timestamp / 1000, tz=timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')
        except Exception as e:
            logger.error(f"Error extracting Twitter date: {e}")
            return None

    # LinkedIn URL Format: https://www.linkedin.com/feed/update/urn:li:activity:7138259296076619777/
    linkedin_match = re.search(r'activity[:-](\d+)', url)

    if linkedin_match:
        try:
            # Extract LinkedIn post ID
            post_id = int(linkedin_match.group(1))
            
            # LinkedIn timestamps are stored in the ID, shifted right by 22 bits
            timestamp_ms = post_id >> 22  # Shift right to extract timestamp in milliseconds
            
            # Convert milliseconds to seconds and format as UTC
            utc_date = datetime.datetime.fromtimestamp(timestamp_ms / 1000, tz=timezone.utc)
            
            return utc_date.strftime('%Y-%m-%d %H:%M:%S UTC')
        except Exception as e:
            logger.error(f"Error extracting LinkedIn date: {e}")
            return None  # If extraction fails, return None

    return None  # Return None for Facebook, Instagram, etc.

def identify_social_media_source(url: str) -> Optional[str]:
    """
    Identifies the social media platform from a URL.
    
    Args:
        url (str): URL to identify
        
    Returns:
        Optional[str]: Name of the platform or None if not identified
    """
    platform_patterns = [
        ("LinkedIn", r"linkedin\.com/posts/|linkedin\.com/pulse/|linkedin\.com/.*/activity-\d+"),
        ("Facebook", r"facebook\.com/.*/photos/\d+"),                     
        ("Pinterest", r"pinterest\.com/.*/pin/\d+"),                      
        ("Instagram", r"instagram\.com/.*/p/[a-zA-Z0-9]+"),               
        ("Twitter/X", r"(?:twitter|x)\.com/.*/status/\d+")
    ]
    
    for platform, pattern in platform_patterns:
        if re.search(pattern, url, re.IGNORECASE):
            return platform
            
    return None

def get_excel_column_widths(file_path: Path) -> Dict[int, float]:
    """
    Read column widths from an existing Excel file.
    
    Args:
        file_path (Path): Path to the Excel file
        
    Returns:
        Dict[int, float]: Dictionary mapping column index to width
    """
    if not file_path.exists():
        return {}
        
    try:
        import openpyxl
        
        # Open the workbook
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        
        # Extract column widths
        col_widths = {}
        for i, column in enumerate(sheet.columns):
            col_letter = openpyxl.utils.get_column_letter(i + 1)
            if hasattr(sheet.column_dimensions[col_letter], 'width'):
                col_widths[i] = sheet.column_dimensions[col_letter].width
        
        return col_widths
    except Exception as e:
        logger.warning(f"⚠️ Could not read column widths from {file_path}: {e}")
        return {}

def save_excel_with_retry(df: pd.DataFrame, file_path: Path, max_retries: int = 5, retry_delay: int = 5) -> bool:
    """
    Tries to save an Excel file, retrying if the file is open.
    Preserves column widths and custom columns from the original file.
    
    Args:
        df (pd.DataFrame): DataFrame to save
        file_path (Path): Path to save the file
        max_retries (int): Maximum number of retry attempts
        retry_delay (int): Delay in seconds between retries
        
    Returns:
        bool: True if successful, False otherwise
    """
    # First, read the column widths from the existing file if it exists
    column_widths = get_excel_column_widths(file_path)
    
    # Check if the file exists to potentially read custom columns
    existing_df = None
    if file_path.exists():
        try:
            existing_df = pd.read_excel(file_path)
        except Exception as e:
            logger.warning(f"⚠️ Could not read existing file for custom columns: {e}")
    
    # If we have an existing DataFrame, preserve any custom columns
    if existing_df is not None:
        # Find columns in the existing file that aren't in our DataFrame
        custom_cols = [col for col in existing_df.columns if col not in df.columns]
        
        if custom_cols:
            logger.info(f"📊 Preserving {len(custom_cols)} custom columns: {', '.join(custom_cols)}")
            
            # For each custom column, copy it over if possible
            for col in custom_cols:
                # Create a mapping of rows in the existing file to rows in our DataFrame
                # This is based on some key that should be unique in both DataFrames
                # For metadata, we can use 'filename'
                # For results, we can use a combination of 'local_image' and 'found_link'
                # For API history, we can use a combination of 'search_type', 'local_image', and 'search_date'
                
                # Determine which key(s) to use based on the columns in our DataFrame
                if 'filename' in df.columns:
                    # This is likely the metadata file
                    key_cols = ['filename']
                elif 'local_image' in df.columns and 'found_link' in df.columns:
                    # This is likely the results file
                    key_cols = ['local_image', 'found_link']
                elif 'search_type' in df.columns and 'local_image' in df.columns and 'search_date' in df.columns:
                    # This is likely the API history file
                    key_cols = ['search_type', 'local_image', 'search_date']
                else:
                    # We don't know how to map rows, so skip this column
                    logger.warning(f"⚠️ Could not determine key columns for {file_path}, skipping custom column '{col}'")
                    continue
                
                # Check if all key columns exist in both DataFrames
                if not all(key in existing_df.columns for key in key_cols):
                    logger.warning(f"⚠️ Key columns {key_cols} not found in existing file, skipping custom column '{col}'")
                    continue
                
                # Add the custom column to our DataFrame, initialized with NaN
                df[col] = np.nan
                
                # Copy values from the existing DataFrame to our DataFrame where keys match
                for _, row in existing_df.iterrows():
                    # Create a filter for matching rows in our DataFrame
                    filter_expr = True
                    for key in key_cols:
                        filter_expr = filter_expr & (df[key] == row[key])
                    
                    # Apply the value from the existing DataFrame to matching rows in our DataFrame
                    if filter_expr.any():
                        df.loc[filter_expr, col] = row[col]
    
    # Now try to save the file with retries
    for attempt in range(max_retries):
        try:
            # Save the DataFrame to Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
                
                # Apply column widths if we have them
                if column_widths:
                    # Access the worksheet
                    worksheet = writer.sheets['Sheet1']
                    
                    # Set column widths
                    for col_idx, width in column_widths.items():
                        if col_idx < len(df.columns):  # Only set width for columns that exist
                            col_letter = openpyxl.utils.get_column_letter(col_idx + 1)
                            worksheet.column_dimensions[col_letter].width = width
            
            logger.info(f"✅ Saved file with preserved formatting: {file_path}")
            return True
        except PermissionError:
            if attempt < max_retries - 1:
                logger.warning(f"❌ Excel file '{file_path}' is open. Please close it. Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
            else:
                logger.error(f"❌ Failed to save after {max_retries} attempts: {file_path}")
                return False
        except Exception as e:
            logger.error(f"❌ Error saving file {file_path}: {str(e)}")
            return False
    
    return False

def is_recently_processed(last_processed_date: str, linkedin_count: int = 0) -> bool:
    """
    Checks if the image was processed recently based on LinkedIn count.
    
    Args:
        last_processed_date (str): Date string when the image was last processed
        linkedin_count (int): Number of LinkedIn links found for this image
        
    Returns:
        bool: True if image was processed recently, False otherwise
    """
    try:
        # Determine the appropriate threshold based on LinkedIn count
        if linkedin_count >= LINKEDIN_HIGH_THRESHOLD:
            threshold_days = HIGH_LINKEDIN_THRESHOLD_DAYS
            logger.debug(f"Using HIGH LinkedIn threshold ({threshold_days} days) for image with {linkedin_count} LinkedIn links")
        else:
            threshold_days = STANDARD_THRESHOLD_DAYS
            logger.debug(f"Using standard threshold ({threshold_days} days) for image with {linkedin_count} LinkedIn links")
        
        # Parse the date string
        last_date = datetime.datetime.strptime(last_processed_date, '%Y-%m-%d %H:%M:%S')
        # Calculate time difference
        time_diff = datetime.datetime.now() - last_date
        # Return True if within threshold
        return time_diff.days < threshold_days
    except (ValueError, TypeError):
        # If date can't be parsed, assume not recently processed
        return False

def update_metadata_counts(metadata_df: pd.DataFrame, results_df: pd.DataFrame) -> pd.DataFrame:
    """
    Update link count columns in the metadata DataFrame based on the results database.
    
    Args:
        metadata_df (pd.DataFrame): Metadata DataFrame to update
        results_df (pd.DataFrame): Results database with found links
        
    Returns:
        pd.DataFrame: Updated metadata DataFrame with count columns
    """
    logger.info("📊 Updating link count statistics in metadata")
    
    # Skip if results database is empty
    if results_df.empty:
        logger.info("📄 Results database is empty, skipping metadata count updates")
        return metadata_df
    
    # Process each image in metadata
    for index, row in metadata_df.iterrows():
        filename = row['filename']
        
        # Get all results for this image
        image_results = results_df[results_df['local_image'] == filename]
        
        if image_results.empty:
            # Reset counts to 0 for images with no results
            metadata_df.loc[index, 'total_links'] = 0
            metadata_df.loc[index, 'exact_match_count'] = 0
            metadata_df.loc[index, 'similar_match_count'] = 0
            metadata_df.loc[index, 'linkedin_count'] = 0
            metadata_df.loc[index, 'instagram_count'] = 0
            metadata_df.loc[index, 'twitter_count'] = 0
            metadata_df.loc[index, 'facebook_count'] = 0
            metadata_df.loc[index, 'pinterest_count'] = 0
            continue
        
        # Calculate total links
        metadata_df.loc[index, 'total_links'] = len(image_results)
        
        # Calculate match type counts
        metadata_df.loc[index, 'exact_match_count'] = len(image_results[image_results['match_type'] == 'Exact Match'])
        metadata_df.loc[index, 'similar_match_count'] = len(image_results[image_results['match_type'] == 'Similar Match'])
        
        # Calculate social media platform counts
        metadata_df.loc[index, 'linkedin_count'] = len(image_results[image_results['source'] == 'LinkedIn'])
        metadata_df.loc[index, 'instagram_count'] = len(image_results[image_results['source'] == 'Instagram'])
        metadata_df.loc[index, 'twitter_count'] = len(image_results[image_results['source'] == 'Twitter/X'])
        metadata_df.loc[index, 'facebook_count'] = len(image_results[image_results['source'] == 'Facebook'])
        metadata_df.loc[index, 'pinterest_count'] = len(image_results[image_results['source'] == 'Pinterest'])
    
    logger.info(f"✅ Updated link counts for {len(metadata_df)} images in metadata")
    return metadata_df

def process_images_and_search() -> None:
    """
    Main function to process images and perform searches with metadata tracking.
    """
    try:
        # Setup folders and files
        images_folder, metadata_path, results_path, api_history_path = setup_folders_and_files()
        
        # Show active search types
        search_types = []
        if DO_EXACT_SEARCH:
            search_types.append("exact matches")
        if DO_SIMILAR_SEARCH:
            search_types.append("similar matches")
            
        if not search_types:
            logger.error("❌ Both DO_EXACT_SEARCH and DO_SIMILAR_SEARCH are set to False. Nothing to do.")
            return
            
        logger.info(f"🔍 Search types enabled: {', '.join(search_types)}")
        
        # Load metadata, results database, and API history database
        metadata_df = load_metadata(metadata_path)
        results_df = load_results_database(results_path)
        api_history_df = load_api_history_database(api_history_path)
        
        # Get current timestamp for tracking
        current_time = datetime.datetime.now()
        current_time_str = current_time.strftime('%Y-%m-%d %H:%M:%S')
        
        # Collect valid image paths
        image_paths = sorted([p for p in images_folder.glob("*.*") if p.suffix.lower() in ALLOWED_EXT])

        if not image_paths:
            logger.error(f"❌ No valid images found in '{images_folder}'.")
            return

        logger.info(f"📂 Found {len(image_paths)} images in '{images_folder}'")
        
        # Track new results for this run
        new_results = []
        metadata_updated = False
        api_history_updated = False
        
        # Count skipped and processed images
        skipped_images = 0
        processed_images = 0
        
        # Track unique and duplicate URLs for summary
        unique_urls_found = 0
        duplicate_urls_skipped = 0
        
        # Process each image
        for img_path in image_paths:
            img_filename = img_path.name
            
            # Check if image exists in metadata
            img_metadata = metadata_df[metadata_df['filename'] == img_filename]
            
            if not img_metadata.empty:
                # Check if image was processed recently based on LinkedIn count
                last_processed = img_metadata.iloc[0]['last_processed_date']
                linkedin_count = img_metadata.iloc[0].get('linkedin_count', 0)  # Default to 0 if column doesn't exist
                
                # Log the processing threshold being used
                if linkedin_count >= LINKEDIN_HIGH_THRESHOLD:
                    logger.info(f"📄 Image {img_filename} has {linkedin_count} LinkedIn links (≥{LINKEDIN_HIGH_THRESHOLD})")
                    logger.info(f"📄 Using {HIGH_LINKEDIN_THRESHOLD_DAYS} day threshold instead of standard {STANDARD_THRESHOLD_DAYS} days")
                
                if is_recently_processed(last_processed, linkedin_count):
                    skipped_images += 1
                    if linkedin_count >= LINKEDIN_HIGH_THRESHOLD:
                        # Always log high-LinkedIn images being skipped
                        logger.info(f"⏭️  Skipped high-LinkedIn image {img_filename} processed within the last {HIGH_LINKEDIN_THRESHOLD_DAYS} days")
                    elif skipped_images % 10 == 0:  # Log only every 10 regular skipped images
                        logger.info(f"⏭️  Skipped {skipped_images} images processed within their threshold periods")
                    continue
                
                # Image already processed before, but not recently, use existing Imgur URL
                img_url = img_metadata.iloc[0]['imgur_url']
                logger.info(f"♻️  Processing: {img_filename} (using existing Imgur URL)")
                
                # Update last processed date
                metadata_df.loc[metadata_df['filename'] == img_filename, 'last_processed_date'] = current_time_str
                metadata_updated = True
            else:
                # New image, upload to Imgur
                logger.info(f"🚀 Processing: {img_filename} (new upload to Imgur)")
                img_url = upload_to_imgur(img_path)
                
                if not img_url:
                    logger.error(f"❌ Skipping {img_filename}, failed to upload.")
                    continue
                
                logger.info(f"✅ Image uploaded: {img_url}")
                
                # Add to metadata
                new_metadata = pd.DataFrame({
                    'filename': [img_filename],
                    'last_processed_date': [current_time_str],
                    'imgur_url': [img_url]
                })
                metadata_df = pd.concat([metadata_df, new_metadata], ignore_index=True)
                metadata_updated = True
            
            processed_images += 1
            
            # Get existing URLs for this image to avoid duplicates
            if not results_df.empty:
                existing_urls = set(results_df[results_df['local_image'] == img_filename]['found_link'].tolist())
            else:
                existing_urls = set()
            
            # Initialize counters for this image
            new_urls_count = 0
            skipped_urls_count = 0
            
            # Process Exact Matches if enabled
            if DO_EXACT_SEARCH:
                logger.info(f"🔍 Performing exact matches search for {img_filename}")
                exact_data, api_history_df = search_google_lens(img_url, img_filename, api_history_df, "exact_matches")
                api_history_updated = True
                
                if exact_data:
                    exact_matches = exact_data.get("exact_matches", [])
                    
                    for index, res in enumerate(exact_matches, start=1):
                        link = res.get("link", "")
                        
                        # Skip if this URL already exists for this image
                        if link in existing_urls:
                            skipped_urls_count += 1
                            duplicate_urls_skipped += 1
                            continue
                        
                        new_urls_count += 1
                        unique_urls_found += 1
                        
                        title = res.get("title", "")
                        post_date = extract_post_date(link)
                        source = identify_social_media_source(link)
                        
                        new_results.append({
                            "local_image": img_filename,
                            "uploaded_url": img_url,
                            "found_link": link,
                            "title": title,
                            "duplicate": 0,  # Initialize as non-duplicate
                            "match_type": "Exact Match",
                            "source": source,
                            "post_date": post_date,
                            "search_date": current_time_str,
                            "order": index  # Add order of appearance
                        })
                else:
                    logger.warning(f"⚠️ No exact matches data received for {img_filename}")
            
            # Process Similar/Visual Matches if enabled
            if DO_SIMILAR_SEARCH:
                logger.info(f"🔍 Performing visual matches search for {img_filename}")
                visual_data, api_history_df = search_google_lens(img_url, img_filename, api_history_df, "visual_matches")
                api_history_updated = True
                
                if visual_data:
                    visual_matches = visual_data.get("visual_matches", [])
                    start_index = 1 if not DO_EXACT_SEARCH else len(exact_matches) + 1 if 'exact_matches' in locals() else 1
                    
                    for index, res in enumerate(visual_matches, start=start_index):
                        link = res.get("link", "")
                        
                        # Skip if this URL already exists for this image
                        if link in existing_urls:
                            skipped_urls_count += 1
                            duplicate_urls_skipped += 1
                            continue
                        
                        new_urls_count += 1
                        unique_urls_found += 1
                        
                        title = res.get("title", "")
                        post_date = extract_post_date(link)
                        source = identify_social_media_source(link)
                        
                        new_results.append({
                            "local_image": img_filename,
                            "uploaded_url": img_url,
                            "found_link": link,
                            "title": title,
                            "duplicate": 0,  # Initialize as non-duplicate
                            "match_type": "Similar Match",
                            "source": source,
                            "post_date": post_date,
                            "search_date": current_time_str,
                            "order": index  # Add order of appearance
                        })
                else:
                    logger.warning(f"⚠️ No visual matches data received for {img_filename}")
            
            # Log summary for this image
            if skipped_urls_count > 0:
                logger.info(f"⏩ {img_filename}: Added {new_urls_count} new URLs, skipped {skipped_urls_count} existing URLs")
            else:
                logger.info(f"✅ {img_filename}: Added {new_urls_count} new URLs")
        
        # Add new results to the database
        if new_results:
            logger.info(f"📊 Adding {len(new_results)} new results to database")
            new_results_df = pd.DataFrame(new_results)
            
            if results_df.empty:
                results_df = new_results_df
            else:
                # Merge new results with existing database, preserving user-added columns
                # First, get all columns from existing results that aren't in new results
                existing_cols = [col for col in results_df.columns if col not in new_results_df.columns]
                
                # For these columns, fill with NaN in the new results
                for col in existing_cols:
                    new_results_df[col] = np.nan
                
                # Concatenate the DataFrames
                results_df = pd.concat([results_df, new_results_df], ignore_index=True)
        else:
            logger.info("📄 No new results to add to database")
        
        # Update the metadata counts based on the results database
        logger.info("🔢 Updating metadata link counts")
        metadata_df = update_metadata_counts(metadata_df, results_df)
        metadata_updated = True
        
        # Mark duplicate URLs in the results database
        logger.info("🔍 Checking for duplicate URLs in results database")
        results_df = mark_duplicate_urls(results_df)
        
        # Save metadata 
        if metadata_updated:
            logger.info("💾 Saving updated metadata with link counts")
            save_excel_with_retry(metadata_df, metadata_path)
        
        # Save results database
        logger.info("💾 Saving updated results database")
        save_excel_with_retry(results_df, results_path)
        
        # Save API history database if updated
        if api_history_updated:
            logger.info("💾 Saving updated API history database")
            save_excel_with_retry(api_history_df, api_history_path)
        
        # Log summary statistics
        logger.info(f"📊 SUMMARY: Processed {processed_images} images, skipped {skipped_images} images")
        logger.info(f"📊 SUMMARY: Found {unique_urls_found} new URLs, skipped {duplicate_urls_skipped} existing URLs")
        logger.info(f"📊 SUMMARY: API history database updated with new API calls")
        logger.info(f"📊 SUMMARY: Metadata updated with link count statistics")
        logger.info(f"📊 SUMMARY: Using two processing thresholds - {HIGH_LINKEDIN_THRESHOLD_DAYS} days for high-LinkedIn images (≥{LINKEDIN_HIGH_THRESHOLD}), {STANDARD_THRESHOLD_DAYS} days for others")
        logger.info(f"✅ Process completed.")
        logger.info(f"💾 Excel files saved in: {METADATA_FOLDER}")
        logger.info(f"🖼️  Images folder: {IMAGES_FOLDER}")
        
    except Exception as e:
        logger.error(f"❌ Error in main process: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())

def mark_duplicate_urls(results_df: pd.DataFrame) -> pd.DataFrame:
    """
    Mark duplicate URLs in the results database.
    
    For each URL:
    - If it appears only once, mark it as 0 (non-duplicate)
    - If it appears multiple times:
      - Mark the oldest entry (by search_date, then post_date, then order) as 2 (primary entry)
      - Mark all other entries as 1 (secondary duplicates)
    
    Args:
        results_df (pd.DataFrame): Results database
        
    Returns:
        pd.DataFrame: Updated results database with duplicate marks
    """
    if results_df.empty:
        logger.info("📄 Results database is empty, skipping duplicate check")
        return results_df
    
    # Initialize duplicate column with 0 (default: non-duplicate)
    results_df['duplicate'] = 0
    
    # Convert date columns to datetime for proper comparison
    # Handle potential NaN values in post_date
    results_df['search_date_dt'] = pd.to_datetime(results_df['search_date'], errors='coerce')
    results_df['post_date_dt'] = pd.to_datetime(results_df['post_date'], errors='coerce')
    
    # Fill NaN post dates with a future date to ensure they're considered last in sorting
    future_date = pd.Timestamp.max
    results_df['post_date_dt'] = results_df['post_date_dt'].fillna(future_date)
    
    # Get counts of each URL
    url_counts = results_df['found_link'].value_counts()
    
    # Get duplicate URLs (appearing more than once)
    duplicate_urls = url_counts[url_counts > 1].index.tolist()
    
    primary_count = 0
    secondary_count = 0
    
    # Process each duplicate URL
    for url in duplicate_urls:
        # Get all rows with this URL
        url_rows = results_df[results_df['found_link'] == url].copy()
        
        # Sort by search_date (oldest first), then post_date (oldest first), then order (lowest first)
        url_rows = url_rows.sort_values(
            by=['search_date_dt', 'post_date_dt', 'order'],
            ascending=[True, True, True]
        )
        
        # Get the index of the first (oldest) row - this one we'll mark as 2 (primary)
        primary_index = url_rows.index[0]
        results_df.loc[primary_index, 'duplicate'] = 2
        primary_count += 1
        
        # Mark all other rows as 1 (secondary duplicates)
        secondary_indices = url_rows.index[1:]
        results_df.loc[secondary_indices, 'duplicate'] = 1
        secondary_count += len(secondary_indices)
    
    # Drop temporary columns
    results_df = results_df.drop(columns=['search_date_dt', 'post_date_dt'])
    
    # Count non-duplicates (entries with value 0)
    non_duplicate_count = (results_df['duplicate'] == 0).sum()
    
    logger.info(f"✅ Duplicate marking complete:")
    logger.info(f"   - {non_duplicate_count} unique URLs (marked as 0)")
    logger.info(f"   - {primary_count} primary duplicate entries (marked as 2)")
    logger.info(f"   - {secondary_count} secondary duplicate entries (marked as 1)")
    logger.info(f"   - {len(results_df)} total entries processed")
    
    return results_df

if __name__ == "__main__":
    # Add missing imports
    import urllib.parse
    import openpyxl
    
    # Parse command line arguments
    args = parse_arguments()
    
    # Initialize global variables
    initialize_globals(args)
    
    logger.info("🚀  Starting Enhanced Image Search Process with API History Tracking")
    logger.info("📊  Custom columns and column widths will be preserved in Excel files")
    logger.info(f"📁  Configuration loaded from: {args.config if args.config else 'linkedin_search_image.json'}")
    process_images_and_search()

"""
Example usage:

1. Using defaults from linkedin_search_image.json:
python linkedin_search_image.py

2. Using a different configuration file:
python linkedin_search_image.py --config "path/to/custom_config.json"

3. Overriding specific settings from command line:
python linkedin_search_image.py --images-folder "C:\\custom\\images" --standard-days 20

4. Enabling similar matches search (overriding config):
python linkedin_search_image.py --enable-similar-search

5. Full example with multiple overrides:
python linkedin_search_image.py \
    --config "custom_config.json" \
    --images-folder "C:\\Users\\rober\\custom_images" \
    --linkedin-high-threshold 30 \
    --enable-similar-search

Note: Command line arguments always take precedence over configuration file values.
"""