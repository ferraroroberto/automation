#!/usr/bin/env python3
"""
LinkedIn Profile Opener - Refactored version
Opens LinkedIn profiles from Excel data with filtering options.
"""

import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Dict, Any

import pandas as pd
import webbrowser
import easygui


def setup_logging(verbose: bool = False) -> logging.Logger:
    """Setup logging configuration."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.StreamHandler()]
    )
    return logging.getLogger(__name__)


def load_config(config_path: str) -> Dict[str, Any]:
    """Load configuration from JSON file."""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        logger.error(f"Configuration file not found: {config_path}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON in configuration file: {e}")
        sys.exit(1)


def update_linkedin_url(url: str, verbose: bool = False) -> str:
    """Update LinkedIn URLs based on the presence of '/company/'."""
    if "/company/" in url:
        url_fix = url + "posts/?feedView=all"
    else:
        url_fix = url + "recent-activity/shares/"
    
    if verbose:
        logger.debug(f"Updated LinkedIn URL: {url_fix}")
    return url_fix


def load_excel_data(file_path: str) -> pd.DataFrame:
    """Load and validate Excel data."""
    logger.info(f"Reading data from: {file_path}")
    
    try:
        df = pd.read_excel(file_path)
        logger.info(f"Successfully loaded {len(df)} rows from Excel file")
        return df
    except PermissionError:
        easygui.msgbox(
            f"Permission denied: Unable to access the file at {file_path}.",
            title="File Access Error"
        )
        logger.error(f"Permission denied accessing file: {file_path}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Error reading Excel file: {e}")
        sys.exit(1)


def get_user_filters(config: Dict[str, Any]) -> Dict[str, str]:
    """Get user filter preferences through GUI."""
    filter_options = config.get('filter_options', {})
    
    filters = {}
    
    # Get alert filter
    filters['alert'] = easygui.choicebox(
        'Select filter for IND_ALERT',
        choices=filter_options.get('alert_choices', ['True', 'False', 'Any'])
    )
    
    # Get star filter
    filters['star'] = easygui.choicebox(
        'Select filter for IND_STAR',
        choices=filter_options.get('star_choices', ['True', 'False', 'Any'])
    )
    
    # Get circle filter
    filters['circle'] = easygui.choicebox(
        'Select circle option',
        choices=filter_options.get('circle_choices', ['Any', 'Top5', 'Key50', 'Vital100', 'None'])
    )
    
    # Get topic filter
    filters['topic'] = easygui.choicebox(
        'Select filter for FK_TOPIC',
        choices=filter_options.get('topic_choices', ['all', 'innovation', 'personal development', 'leadership and management', 'LinkedIn', 'visual illustration'])
    )
    
    return filters


def apply_filters(df: pd.DataFrame, filters: Dict[str, str]) -> pd.DataFrame:
    """Apply user filters to the dataframe."""
    logger.info("Applying filters to data...")
    
    # Start with basic mask (URL_LINKEDIN not null)
    mask = df['URL_LINKEDIN'].notnull()
    
    # Apply alert filter
    if filters['alert'] != 'Any':
        mask &= (df['IND_ALERT'] == (filters['alert'] == 'True'))
    
    # Apply star filter
    if filters['star'] != 'Any':
        mask &= (df['IND_STAR'] == (filters['star'] == 'True'))
    
    # Apply circle filter
    if filters['circle'] == 'None':
        mask &= df['FK_CIRCLE'].isnull()
    elif filters['circle'] != 'Any':
        mask &= (df['FK_CIRCLE'] == filters['circle'])
    
    # Apply topic filter
    if filters['topic'] != 'all':
        mask &= (df['FK_TOPIC'] == filters['topic'])
    
    filtered_df = df.loc[mask].copy()
    logger.info(f"Filtered data. Number of rows: {len(filtered_df)}")
    
    return filtered_df


def update_urls(df: pd.DataFrame, verbose: bool = False) -> pd.DataFrame:
    """Update LinkedIn URLs in the dataframe."""
    logger.info("Updating LinkedIn URLs...")
    df['URL_LINKEDIN'] = df['URL_LINKEDIN'].apply(lambda url: update_linkedin_url(url, verbose))
    return df


def open_linkedin_profiles(df: pd.DataFrame, chunk_size: int) -> None:
    """Open LinkedIn profiles in browser with chunking."""
    total_urls = len(df)
    if total_urls == 0:
        logger.warning("No profiles to open after filtering")
        return
    
    logger.info(f"Opening {total_urls} LinkedIn profiles in chunks of {chunk_size}")
    num_urls_opened = 0
    
    for i in range(0, total_urls, chunk_size):
        chunk = df.iloc[i:i+chunk_size]
        
        for j, (_, row) in enumerate(chunk.iterrows(), start=1):
            num_urls_opened += 1
            url = row['URL_LINKEDIN']
            topic = row['FK_TOPIC']
            person = row['DE_PERSON']
            
            logger.info(f"Opening URL {num_urls_opened:03} ({topic} - {person})")
            webbrowser.open(url)
        
        # Ask for continuation if there are more URLs
        if num_urls_opened < total_urls:
            input(f"Press enter to open the next {chunk_size} URLs")
    
    logger.info("Process finished.")


def main():
    """Main execution function."""
    # Get the directory where this script is located
    script_dir = Path(__file__).parent
    
    parser = argparse.ArgumentParser(
        description="Open LinkedIn profiles from Excel data with filtering options",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python linkedin_open.py
  python linkedin_open.py --config custom_config.json
  python linkedin_open.py --verbose --chunk-size 5
        """
    )
    
    parser.add_argument(
        '--config',
        default=str(script_dir / 'linkedin_open.json'),
        help='Path to JSON configuration file (default: linkedin_open.json in script directory)'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose logging'
    )
    
    parser.add_argument(
        '--chunk-size',
        type=int,
        help='Number of profiles to open at once (overrides config)'
    )
    
    args = parser.parse_args()
    
    # Setup logging
    global logger
    logger = setup_logging(args.verbose)
    
    # Load configuration
    config = load_config(args.config)
    logger.info(f"Loaded configuration from: {args.config}")
    
    # Override config with command line arguments
    if args.chunk_size:
        config['default_chunk_size'] = args.chunk_size
    
    # Load Excel data
    df = load_excel_data(config['file_path'])
    
    # Sort data
    df = df.sort_values(by=['FK_TOPIC', 'DE_PERSON'])
    logger.info("Data sorted by topic and person")
    
    # Get user filters
    filters = get_user_filters(config)
    
    # Apply filters
    df = apply_filters(df, filters)
    
    # Update URLs
    df = update_urls(df, config.get('verbose', False))
    
    # Get chunk size from user or config
    if args.chunk_size is None:
        num_profiles = input(f"Enter the number of profiles to open (default is {config['default_chunk_size']}): ")
        try:
            chunk_size = int(num_profiles) if num_profiles.strip() else config['default_chunk_size']
        except ValueError:
            chunk_size = config['default_chunk_size']
    else:
        chunk_size = args.chunk_size
    
    # Open profiles
    open_linkedin_profiles(df, chunk_size)


if __name__ == "__main__":
    main()
