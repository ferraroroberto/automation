"""
LinkedIn Profiles Data Orchestrator Module - Chrome DevTools Version

Orchestrates the LinkedIn profile data extraction and Excel file merging using DevTools.
Extracts data from Chrome tabs and merges new records into an existing Excel file.
"""

import logging
import json
import os
import threading
from typing import List, Dict, Tuple
import pandas as pd
from pynput import keyboard

# Import the DevTools extractor module
from linkedin_profiles_data_extractor_devtools import LinkedInProfileExtractorDevTools, ChromeDevToolsClient
from excel_format_manager import (
    save_excel_format_to_json,
    apply_format_from_json,
    convert_url_columns_to_hyperlinks
)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class LinkedInProfilesDataOrchestratorDevTools:
    """Orchestrates LinkedIn profile data extraction and Excel file merging using DevTools."""

    def __init__(self, config_path: str, debug_port: int = 9222):
        """Initialize orchestrator with configuration.

        Args:
            config_path: Path to the JSON configuration file.
            debug_port: Chrome remote debugging port
        """
        self.config = self.load_config(config_path)
        self.extractor = LinkedInProfileExtractorDevTools(debug_port)
        self.exit_requested = False  # Flag to track if 'x' key was pressed
        self.keyboard_listener = None  # Keyboard listener for 'x' keypress

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

    def wait_for_file_access(self, file_path: str) -> bool:
        """Wait for user to close the file if it's open.

        Args:
            file_path: Path to the file that might be open.

        Returns:
            True if user wants to continue, False if they want to exit.
        """
        while not self.check_file_accessible(file_path):
            print(f"\n⚠️  The file '{os.path.basename(file_path)}' appears to be open in another application.")
            print("Please close the file and press ENTER to continue, or type 'x' to exit:")
            user_input = input().strip().lower()
            if user_input == 'x':
                return False
        return True

    def merge_to_excel(self, new_profiles: List[Dict[str, str]], destination_file: str) -> Tuple[int, int]:
        """Merge new profiles into existing Excel file based on name column.

        Args:
            new_profiles: List of new profile dictionaries to merge.
            destination_file: Path to destination Excel file.

        Returns:
            Tuple of (new_records_added, total_records_after_merge)
        """
        if not new_profiles:
            logger.info("ℹ️  No new profiles to merge")
            return 0, 0

        try:
            # Check if destination file exists
            if os.path.exists(destination_file):
                # Read existing data
                existing_df = pd.read_excel(destination_file, engine='openpyxl')
                logger.info(f"📖 Read existing Excel file with {len(existing_df)} rows")

                # Convert to lowercase for case-insensitive comparison
                existing_names = set()
                if 'name' in existing_df.columns:
                    existing_names = set(
                        str(name).lower().strip()
                        for name in existing_df['name'].dropna()
                        if str(name).strip()
                    )

                # Filter new profiles to exclude duplicates
                filtered_new_profiles = []
                duplicates_found = 0

                for profile in new_profiles:
                    profile_name = profile.get('name', '').lower().strip()
                    if profile_name and profile_name not in existing_names:
                        filtered_new_profiles.append(profile)
                        existing_names.add(profile_name)
                    else:
                        duplicates_found += 1

                if duplicates_found > 0:
                    logger.info(f"⏭️  Skipped {duplicates_found} duplicate profiles")

                if not filtered_new_profiles:
                    logger.info("ℹ️  All new profiles were duplicates - nothing to merge")
                    return 0, len(existing_df)

                # Append new profiles to existing data
                new_df = pd.DataFrame(filtered_new_profiles)
                combined_df = pd.concat([existing_df, new_df], ignore_index=True)

                logger.info(f"📊 Adding {len(filtered_new_profiles)} new profiles to existing {len(existing_df)} profiles")

            else:
                # Create new file with all profiles
                combined_df = pd.DataFrame(new_profiles)
                filtered_new_profiles = new_profiles
                logger.info(f"📄 Creating new Excel file with {len(new_profiles)} profiles")

            # Save to Excel with formatting
            with pd.ExcelWriter(destination_file, engine='openpyxl') as writer:
                combined_df.to_excel(writer, sheet_name='Profiles', index=False)

                # Apply formatting if format file exists
                format_file = self.config.get('excel_format_file')
                if format_file and os.path.exists(format_file):
                    apply_format_from_json(writer, format_file)
                    convert_url_columns_to_hyperlinks(writer, 'Profiles', ['url'])

            logger.info(f"✅ Successfully merged data to {destination_file}")
            return len(filtered_new_profiles), len(combined_df)

        except Exception as e:
            logger.error(f"❌ Failed to merge to Excel: {e}")
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

        try:
            # Setup keyboard listener for exit
            self.setup_keyboard_listener()
            logger.info("⌨️  Press 'x' at any time to stop the extraction")

            # Get configuration values
            destination_file = self.config.get('destination_file')
            max_tabs = self.config.get('max_tabs', 50)

            if not destination_file:
                logger.error("❌ destination_file not specified in configuration")
                return

            # Ensure destination directory exists
            os.makedirs(os.path.dirname(destination_file), exist_ok=True)

            # Wait for file access if it exists
            if os.path.exists(destination_file):
                logger.info(f"🔍 Checking access to destination file: {destination_file}")
                if not self.wait_for_file_access(destination_file):
                    logger.info("⏹️  User chose to exit")
                    return

            # Step 1: Extract profiles from Chrome tabs using DevTools
            logger.info("📋 Step 1: Extracting profiles from Chrome tabs using DevTools...")
            profiles, tabs_processed = self.extractor.extract_all_tabs_devtools(
                max_tabs=max_tabs,
                exit_check=lambda: self.exit_requested
            )

            if not profiles:
                logger.warning("⚠️  No profiles extracted")
                return

            # Step 2: Merge to Excel
            logger.info("📋 Step 2: Merging profiles to Excel file...")

            # Save format before merging (if file exists)
            if os.path.exists(destination_file):
                format_file = self.config.get('excel_format_file')
                if format_file:
                    save_excel_format_to_json(destination_file, format_file)

            new_records, total_records = self.merge_to_excel(profiles, destination_file)

            # Step 3: Report results
            logger.info("📋 Step 3: Extraction and merge completed")
            print("\n" + "="*80)
            print("EXTRACTION RESULTS (DevTools Version)")
            print("="*80)
            print(f"Profiles extracted: {len(profiles)}")
            print(f"Tabs processed: {tabs_processed}")
            print(f"New records added: {new_records}")
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
    # Load configuration from the same directory as this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, "linkedin_profiles_data.json")

    if not os.path.exists(config_path):
        logger.error(f"❌ Configuration file not found: {config_path}")
        logger.error(f"   Current working directory: {os.getcwd()}")
        logger.error(f"   Script directory: {script_dir}")
        return

    # Initialize and run orchestrator
    orchestrator = LinkedInProfilesDataOrchestratorDevTools(config_path)

    # Ask user what they want to do
    print("\nLinkedIn Profiles Data Orchestrator (DevTools Version)")
    print("=" * 55)
    print("1. Run full extraction and merge to Excel")
    print("2. Run quick test (single profile)")
    print("3. Exit")
    print()

    while True:
        try:
            choice = input("Enter your choice (1-3): ").strip()
            if choice == '1':
                orchestrator.run_extraction_and_merge()
                break
            elif choice == '2':
                orchestrator.run_quick_test()
                break
            elif choice == '3':
                print("Exiting...")
                break
            else:
                print("Invalid choice. Please enter 1, 2, or 3.")
        except KeyboardInterrupt:
            print("\nExiting...")
            break


if __name__ == "__main__":
    main()
