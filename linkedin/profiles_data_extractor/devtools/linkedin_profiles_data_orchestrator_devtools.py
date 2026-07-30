"""
LinkedIn Profiles Data Orchestrator Module - Chrome DevTools Version

Orchestrates the LinkedIn profile data extraction and Excel file merging using DevTools.
Extracts data from Chrome tabs and merges new records into an existing Excel file.

Usage:
    python linkedin_profiles_data_orchestrator_devtools.py [--run] [--save-format]
    
Options:
    --run, -r           Run full extraction and merge to Excel (default if no options)
    --save-format, -s   Save Excel format to JSON before extraction
    --test, -t          Run quick test (single profile)
    --help, -h          Show this help message
"""

import argparse
import json
import os
import sys
import time
from typing import List, Dict, Tuple
import pandas as pd
from pynput import keyboard

# Add current directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the DevTools extractor module
from linkedin_profiles_data_extractor_devtools import LinkedInProfileExtractorDevTools, ChromeDevToolsClient
from _chrome_client import setup_logging

# Import from common directory using importlib for reliable path handling
import importlib.util

# Calculate path from devtools directory to common directory (../common)
script_dir = os.path.dirname(os.path.abspath(__file__))
common_dir = os.path.join(script_dir, '..', 'common')
common_dir = os.path.abspath(common_dir)  # Normalize path

# excel_format_manager.py bare-imports its sibling _logging_setup, so common_dir
# must be on sys.path before exec_module() runs it below.
sys.path.insert(0, common_dir)

excel_format_path = os.path.join(common_dir, "excel_format_manager.py")
excel_format_spec = importlib.util.spec_from_file_location("excel_format_manager", excel_format_path)
excel_format_manager = importlib.util.module_from_spec(excel_format_spec)
excel_format_spec.loader.exec_module(excel_format_manager)

# Import the functions we need
save_excel_format_to_json = excel_format_manager.save_excel_format_to_json
apply_format_from_json = excel_format_manager.apply_format_from_json
convert_url_columns_to_hyperlinks = excel_format_manager.convert_url_columns_to_hyperlinks

logger = setup_logging(__name__)


class LinkedInProfilesDataOrchestratorDevTools:
    """Orchestrates LinkedIn profile data extraction and Excel file merging using DevTools."""

    def __init__(self, config_path: str, debug_port: int = 9222, save_format: bool = False):
        """Initialize orchestrator with configuration.

        Args:
            config_path: Path to the JSON configuration file.
            debug_port: Chrome remote debugging port
            save_format: Whether to save Excel format to JSON before extraction
        """
        self.config = self.load_config(config_path)
        self.extractor = LinkedInProfileExtractorDevTools(debug_port)
        self.exit_requested = False  # Flag to track if 'x' key was pressed
        self.keyboard_listener = None  # Keyboard listener for 'x' keypress
        self.save_format = save_format

    def load_config(self, config_path: str) -> Dict:
        """Load configuration from JSON file.

        Args:
            config_path: Path to the configuration file.

        Returns:
            Configuration dictionary.
        """
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            logger.info(f"✅ Configuration loaded from {config_path}")
            return config
        except Exception as e:
            logger.error(f"❌ Failed to load configuration: {e}")
            raise

    def check_file_accessible(self, file_path: str) -> bool:
        """Check if file is accessible for reading/writing.

        Args:
            file_path: Path to the file to check.

        Returns:
            True if file is accessible, False otherwise.
        """
        try:
            with open(file_path, 'r+b') as f:
                pass
            return True
        except (PermissionError, OSError):
            return False

    def wait_for_file_access(self, file_path: str, max_retries: int = 30, retry_delay: float = 1.0) -> bool:
        """Wait for file to become accessible (auto-retry).

        Args:
            file_path: Path to the file that might be open.
            max_retries: Maximum number of retries before giving up.
            retry_delay: Delay in seconds between retries.

        Returns:
            True if file is accessible, False if max retries exceeded.
        """
        retries = 0
        while not self.check_file_accessible(file_path):
            retries += 1
            if retries > max_retries:
                logger.error(f"❌ File '{os.path.basename(file_path)}' is not accessible after {max_retries} retries")
                return False
            logger.warning(f"⚠️  File '{os.path.basename(file_path)}' is open. Retrying in {retry_delay}s... ({retries}/{max_retries})")
            time.sleep(retry_delay)
        return True

    def merge_to_excel(self, new_profiles: List[Dict[str, str]], destination_file: str) -> Tuple[int, int, int]:
        """Merge new profiles into existing Excel file based on name column.

        Args:
            new_profiles: List of new profile dictionaries to merge.
            destination_file: Path to the destination Excel file.

        Returns:
            Tuple of (inserted_count, skipped_count).
        """
        # Columns to use from the profiles (as specified in requirements)
        columns_to_use = ['name', 'url_profile', 'job_title', 'follows_from', 'company', 'location']

        # Filter new profiles to only include specified columns
        filtered_profiles = []
        for profile in new_profiles:
            filtered_profile = {col: profile.get(col, '') for col in columns_to_use}
            # Only include profiles that have a name (required for deduplication)
            if filtered_profile.get('name', '').strip():
                filtered_profiles.append(filtered_profile)

        if not filtered_profiles:
            logger.warning("⚠️  No valid profiles to merge (missing name field)")
            return 0, 0

        # Check if destination file exists
        if not os.path.exists(destination_file):
            logger.info(f"📄 Destination file doesn't exist, creating new file: {destination_file}")
            # Create new file with all profiles
            df_new = pd.DataFrame(filtered_profiles)
            df_new.to_excel(destination_file, index=False, engine='openpyxl')
            logger.info(f"✅ Created new Excel file with {len(filtered_profiles)} profiles")
            return len(filtered_profiles), 0, len(filtered_profiles)

        try:
            # Check if file is accessible before reading
            if not self.wait_for_file_access(destination_file):
                logger.info("ℹ️  User chose to exit - skipping Excel merge")
                return 0, len(filtered_profiles), 0

            # Read existing Excel file
            df_existing = pd.read_excel(destination_file, engine='openpyxl')
            logger.info(f"📖 Read existing Excel file with {len(df_existing)} rows")

            # Get existing names (case-insensitive comparison)
            existing_names = set()
            if 'name' in df_existing.columns:
                existing_names = set(
                    name.lower().strip()
                    for name in df_existing['name'].dropna()
                    if isinstance(name, str) and name.strip()
                )

            # Filter out profiles that already exist
            new_profiles_filtered = []
            skipped_count = 0

            for profile in filtered_profiles:
                profile_name = profile.get('name', '').lower().strip()
                if profile_name and profile_name not in existing_names:
                    new_profiles_filtered.append(profile)
                else:
                    skipped_count += 1
                    if profile_name:
                        logger.info(f"⏭️  Skipped existing profile: {profile['name']}")
                    else:
                        logger.info("⏭️  Skipped profile with missing name")

            # If no new profiles to add, return early
            if not new_profiles_filtered:
                logger.info("ℹ️  No new profiles to add - all profiles already exist")
                return 0, len(filtered_profiles), len(df_existing)

            # Create DataFrame from new profiles
            df_new = pd.DataFrame(new_profiles_filtered)

            # Ensure all required columns exist in both DataFrames
            all_columns = set(df_existing.columns.tolist() + columns_to_use)
            for col in all_columns:
                if col not in df_existing.columns:
                    df_existing[col] = ''
                if col not in df_new.columns:
                    df_new[col] = ''

            # Reorder columns to match the desired order
            df_new = df_new[columns_to_use]

            # Append new profiles to existing data
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)

            # Check file access before writing
            if not self.wait_for_file_access(destination_file):
                logger.info("ℹ️  User chose to exit - changes not saved")
                return 0, len(filtered_profiles), len(df_existing)

            # Save back to Excel
            df_combined.to_excel(destination_file, index=False, engine='openpyxl')

            inserted_count = len(new_profiles_filtered)
            total_records = len(df_combined)
            logger.info(f"✅ Successfully merged {inserted_count} new profiles into {destination_file}")
            logger.info(f"📊 Summary: {inserted_count} inserted, {skipped_count} skipped, {total_records} total")

            return inserted_count, skipped_count, total_records

        except Exception as e:
            logger.error(f"❌ Failed to merge profiles into Excel: {e}")
            raise

    def setup_keyboard_listener(self) -> None:
        """Set up keyboard listener for 'x' key to exit extraction."""
        def on_press(key):
            try:
                if key.char.lower() == 'x':
                    self.exit_requested = True
                    logger.info("🛑 Exit requested by user (x key pressed)")
            except AttributeError:
                pass  # Special keys don't have char attribute

        self.keyboard_listener = keyboard.Listener(on_press=on_press)
        self.keyboard_listener.start()

    def stop_keyboard_listener(self) -> None:
        """Stop the keyboard listener."""
        if self.keyboard_listener:
            self.keyboard_listener.stop()
            self.keyboard_listener = None

    def run_extraction_and_merge(self) -> None:
        """Run the complete extraction and merge process using DevTools."""
        logger.info("🎯 Starting LinkedIn Profiles Data Orchestrator (DevTools Version)")
        logger.info("=" * 80)

        # Get configuration values
        destination_file = self.config.get('destination_file')
        max_tabs = self.config.get('max_tabs', 50)

        if not destination_file:
            logger.error("❌ destination_file not specified in configuration")
            return

        # Ensure destination directory exists
        os.makedirs(os.path.dirname(destination_file), exist_ok=True)

        # Create paths to common directory files (same as in main())
        script_dir = os.path.dirname(os.path.abspath(__file__))
        common_dir = os.path.join(os.path.dirname(script_dir), "common")
        config_path = os.path.join(common_dir, "linkedin_profiles_data.json")
        format_json_path = os.path.join(common_dir, "excel_format_spec.json")

        try:
            # STEP 0: Optionally save format to JSON (controlled by --save-format flag)
            if self.save_format and os.path.exists(destination_file):
                logger.info("")
                logger.info("=" * 80)
                logger.info("STEP 0: Saving Excel format to JSON")
                logger.info("=" * 80)
                logger.info("💾 Saving Excel format to JSON...")
                success = save_excel_format_to_json(destination_file, format_json_path)
                if success:
                    logger.info("✅ Format saved successfully")
                else:
                    logger.warning("⚠️  Failed to save format, continuing anyway...")

            # Setup keyboard listener for exit
            self.setup_keyboard_listener()
            logger.info("")
            logger.info("⌨️  Press 'x' at any time to stop the extraction")

            # Wait for file access if it exists
            if os.path.exists(destination_file):
                logger.info(f"🔍 Checking access to destination file: {destination_file}")
                if not self.wait_for_file_access(destination_file):
                    logger.info("⏹️  User chose to exit")
                    return

            # STEP 1: Extract profiles from Chrome tabs using DevTools
            logger.info("")
            logger.info("=" * 80)
            logger.info("STEP 1: Extracting data from Chrome tabs...")
            logger.info("=" * 80)
            profiles, tabs_processed = self.extractor.extract_all_tabs_devtools(
                max_tabs=max_tabs,
                exit_check=lambda: self.exit_requested
            )

            if not profiles:
                logger.warning("⚠️  No profiles extracted")
                return

            # STEP 2: Merge to destination Excel file
            logger.info("")
            logger.info("=" * 80)
            logger.info("STEP 2: Merging data to destination Excel file...")
            logger.info("=" * 80)

            # Ensure destination directory exists
            destination_dir = os.path.dirname(destination_file)
            os.makedirs(destination_dir, exist_ok=True)

            inserted_count, skipped_count, total_records = self.merge_to_excel(profiles, destination_file)

            # STEP 3: Apply format and convert URLs
            logger.info("")
            logger.info("=" * 80)
            logger.info("STEP 3: Applying format and converting URLs...")
            logger.info("=" * 80)

            if os.path.exists(format_json_path):
                logger.info("📋 Applying format from JSON...")
                success = apply_format_from_json(destination_file, format_json_path)
                if success:
                    logger.info("✅ Format applied successfully")
                else:
                    logger.warning("⚠️  Failed to apply format")
            else:
                logger.warning(f"⚠️  Format JSON not found: {format_json_path} - skipping format application")

            logger.info("")
            logger.info("📎 Converting URL columns to hyperlinks...")
            success = convert_url_columns_to_hyperlinks(destination_file)
            if success:
                logger.info("✅ URLs converted successfully")
            else:
                logger.warning("⚠️  Failed to convert URLs")

            # Display final summary
            logger.info("")
            print("\n" + "="*80)
            print("ORCHESTRATION COMPLETE (DevTools Version)")
            print("="*80)
            print(f"Total profiles extracted: {len(profiles)}")
            print(f"Tabs processed: {tabs_processed}")
            print(f"Profiles inserted: {inserted_count}")
            print(f"Profiles skipped (already exist): {skipped_count}")
            print(f"Total records in file: {total_records}")
            print(f"Destination file: {destination_file}")
            print("="*80 + "\n")

            logger.info("✅ Orchestration completed successfully")

        except KeyboardInterrupt:
            logger.info("⏹️  Orchestration interrupted by user")
        except Exception as e:
            logger.error(f"❌ Orchestration failed: {e}")
            raise
        finally:
            self.stop_keyboard_listener()

    def run_quick_test(self) -> None:
        """Run a quick test to verify DevTools connection and single profile extraction."""
        logger.info("🧪 Running DevTools connection test")

        try:
            if not self.extractor.connect_to_chrome():
                logger.error("❌ DevTools connection test failed")
                return

            # Test single profile extraction
            profile = self.extractor.extract_profile_data()
            logger.info("📊 Test extraction result:")
            for key, value in profile.items():
                logger.info(f"   {key}: {value}")

            logger.info("✅ DevTools test completed successfully")

        finally:
            self.extractor.chrome.close()


def main() -> None:
    """Main execution function."""
    # Parse command-line arguments
    parser = argparse.ArgumentParser(
        description="LinkedIn Profiles Data Orchestrator (DevTools Version)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python linkedin_profiles_data_orchestrator_devtools.py --run
    python linkedin_profiles_data_orchestrator_devtools.py --run --save-format
    python linkedin_profiles_data_orchestrator_devtools.py --test
        """
    )
    parser.add_argument(
        '--run', '-r',
        action='store_true',
        help='Run full extraction and merge to Excel'
    )
    parser.add_argument(
        '--save-format', '-s',
        action='store_true',
        help='Save Excel format to JSON before extraction'
    )
    parser.add_argument(
        '--test', '-t',
        action='store_true',
        help='Run quick test (single profile)'
    )
    
    args = parser.parse_args()
    
    # Default to --run if no action specified
    if not args.run and not args.test:
        args.run = True
    
    # Load configuration from the common directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    common_dir = os.path.join(os.path.dirname(script_dir), "common")
    config_path = os.path.join(common_dir, "linkedin_profiles_data.json")

    if not os.path.exists(config_path):
        logger.error(f"❌ Configuration file not found: {config_path}")
        logger.error(f"   Current working directory: {os.getcwd()}")
        logger.error(f"   Script directory: {script_dir}")
        return

    # Initialize orchestrator with save_format option
    orchestrator = LinkedInProfilesDataOrchestratorDevTools(config_path, save_format=args.save_format)

    try:
        if args.run:
            logger.info("🚀 Running full extraction and merge to Excel...")
            orchestrator.run_extraction_and_merge()
        elif args.test:
            orchestrator.run_quick_test()
    except KeyboardInterrupt:
        logger.info("⏹️  Interrupted by user")
    except Exception as e:
        logger.error(f"❌ An error occurred: {e}")


if __name__ == "__main__":
    main()
