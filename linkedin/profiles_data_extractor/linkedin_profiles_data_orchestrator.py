"""
LinkedIn Profiles Data Orchestrator Module

Orchestrates the LinkedIn profile data extraction and Excel file merging.
Extracts data from Chrome tabs and merges new records into an existing Excel file.
"""

import logging
import json
import os
from typing import List, Dict, Tuple
import pandas as pd

# Import the existing extractor module
from linkedin_profiles_data_extractor import LinkedInProfileExtractor, save_to_excel

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class LinkedInProfilesDataOrchestrator:
    """Orchestrates LinkedIn profile data extraction and Excel file merging."""

    def __init__(self, config_path: str):
        """Initialize orchestrator with configuration.

        Args:
            config_path: Path to the JSON configuration file.
        """
        self.config = self.load_config(config_path)
        self.extractor = LinkedInProfileExtractor()

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

    def merge_to_excel(self, new_profiles: List[Dict[str, str]], destination_file: str) -> Tuple[int, int]:
        """Merge new profiles into existing Excel file based on name column.

        Args:
            new_profiles: List of new profile dictionaries to merge.
            destination_file: Path to the destination Excel file.

        Returns:
            Tuple of (inserted_count, skipped_count).
        """
        # Columns to use from the profiles (as specified in requirements)
        columns_to_use = ['name', 'url', 'job_title', 'follows_from', 'company', 'location']

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
            return len(filtered_profiles), 0

        try:
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
                return 0, len(filtered_profiles)

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

            # Save back to Excel
            df_combined.to_excel(destination_file, index=False, engine='openpyxl')

            inserted_count = len(new_profiles_filtered)
            logger.info(f"✅ Successfully merged {inserted_count} new profiles into {destination_file}")
            logger.info(f"📊 Summary: {inserted_count} inserted, {skipped_count} skipped")

            return inserted_count, skipped_count

        except Exception as e:
            logger.error(f"❌ Failed to merge profiles into Excel: {e}")
            raise

    def run_extraction_cycle(self) -> List[Dict[str, str]]:
        """Run the complete extraction cycle.

        Returns:
            List of extracted profile dictionaries.
        """
        logger.info("🚀 Starting LinkedIn profile data extraction cycle")

        # Extract data from all tabs using the existing extractor
        max_tabs = self.config.get('max_tabs', 50)
        profiles_data = self.extractor.extract_all_tabs(max_tabs)

        if profiles_data:
            logger.info(f"✅ Extracted {len(profiles_data)} profiles from {max_tabs} tabs")
        else:
            logger.warning("⚠️  No profile data was extracted")

        return profiles_data

    def orchestrate(self) -> None:
        """Main orchestration method that runs extraction and merging."""
        logger.info("🎯 Starting LinkedIn Profiles Data Orchestrator")
        logger.info("=" * 80)

        try:
            # Step 1: Run extraction cycle
            logger.info("📋 Step 1: Extracting data from Chrome tabs...")
            profiles_data = self.run_extraction_cycle()

            if not profiles_data:
                logger.warning("⚠️  No data extracted - skipping Excel merge")
                return

            # Step 2: Merge to destination Excel file
            logger.info("📋 Step 2: Merging data to destination Excel file...")

            destination_file = self.config.get('destination_file')
            if not destination_file:
                logger.error("❌ destination_file not specified in configuration")
                return

            # Ensure destination directory exists
            destination_dir = os.path.dirname(destination_file)
            os.makedirs(destination_dir, exist_ok=True)

            # Merge profiles
            inserted_count, skipped_count = self.merge_to_excel(profiles_data, destination_file)

            # Display final summary
            print("\n" + "="*80)
            print("ORCHESTRATION COMPLETE")
            print("="*80)
            print(f"Total profiles extracted: {len(profiles_data)}")
            print(f"Profiles inserted: {inserted_count}")
            print(f"Profiles skipped (already exist): {skipped_count}")
            print(f"Destination file: {destination_file}")
            print("="*80 + "\n")

            logger.info("✅ Orchestration completed successfully")

        except Exception as e:
            logger.error(f"❌ Orchestration failed: {e}")
            raise


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
    orchestrator = LinkedInProfilesDataOrchestrator(config_path)
    orchestrator.orchestrate()


if __name__ == "__main__":
    main()
