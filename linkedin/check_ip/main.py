#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
LinkedIn Image Search - Main Application

This is the main orchestration script for the LinkedIn image search application.
It coordinates all the modules to perform image searches using Google Lens via SerpAPI.

Features:
- Metadata management: Tracks images already processed and their Imgur URLs
- Cumulative database: Preserves historical search results and user annotations
- API history database: Stores complete details of all API calls including raw JSON responses
- Optimization: Avoids re-uploading images that were previously processed
- Social media extraction: Extracts and formats post dates from various platforms
- Time-based processing: Skips images processed within the last N days based on LinkedIn count
- Configurable search types: Control whether to do exact matches or enable similar matches

Author: Roberto (Refactored by Claude)
Date: March 2025
"""

import datetime
import logging
import numpy as np
import pandas as pd
import sys
from pathlib import Path

# Import our modules
from config import parse_arguments, get_app_config
from storage import StorageManager
from imgur_client import ImgurClient
from search_client import SearchClient
from data_processor import DataProcessor

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class LinkedInImageSearchApp:
    """Main application class for LinkedIn image search."""
    
    def __init__(self, config: dict):
        """
        Initialize the application with configuration.
        
        Args:
            config (dict): Application configuration
        """
        self.config = config
        
        # Initialize components
        self.storage_manager = StorageManager(config["metadata_folder"])
        self.imgur_client = ImgurClient(
            config["imgur_client_id"],
            config["imgur_access_token"]
        )
        self.search_client = SearchClient(
            config["serpapi_key"],
            config["search_endpoint"]
        )
        self.data_processor = DataProcessor()
        
        # Validate API credentials
        self._validate_credentials()
    
    def _validate_credentials(self):
        """Validate API credentials."""
        logger.info("🔐 Validating API credentials...")
        
        # Validate Imgur credentials (optional - don't fail if validation fails)
        try:
            self.imgur_client.validate_credentials()
        except Exception as e:
            logger.warning(f"⚠️ Imgur credential validation failed: {e}")
        
        # Validate SerpAPI key
        if not self.search_client.validate_api_key():
            logger.error("❌ SerpAPI validation failed. Please check your API key.")
            sys.exit(1)
    
    def run(self):
        """Run the main image search process."""
        try:
            # Setup folders and files
            images_folder, metadata_path, results_path, api_history_path = self.storage_manager.setup_folders_and_files(
                self.config["images_folder"],
                self.config["metadata_file"],
                self.config["results_file"],
                self.config["api_history_file"]
            )
            
            # Show active search types
            search_types = []
            if self.config["do_exact_search"]:
                search_types.append("exact matches")
            if self.config["do_similar_search"]:
                search_types.append("similar matches")
                
            if not search_types:
                logger.error("❌ Both exact and similar search are disabled. Nothing to do.")
                return
                
            logger.info(f"🔍 Search types enabled: {', '.join(search_types)}")
            
            # Load all databases
            metadata_df = self.storage_manager.load_metadata(metadata_path)
            results_df = self.storage_manager.load_results_database(results_path)
            api_history_df = self.storage_manager.load_api_history_database(api_history_path)
            
            # Get current timestamp for tracking
            current_time = datetime.datetime.now()
            current_time_str = current_time.strftime('%Y-%m-%d %H:%M:%S')
            
            # Collect valid image paths
            image_paths = self.data_processor.collect_valid_image_paths(
                images_folder, 
                self.config["allowed_extensions"]
            )

            if not image_paths:
                logger.error(f"❌ No valid images found in '{images_folder}'.")
                return

            logger.info(f"📂 Found {len(image_paths)} images in '{images_folder}'")
            
            # Process images
            processing_results = self._process_images(
                image_paths, metadata_df, results_df, api_history_df, current_time_str
            )
            
            # Unpack results
            (metadata_df, results_df, api_history_df, new_results, 
             skipped_images, processed_images, unique_urls_found, duplicate_urls_skipped) = processing_results
            
            # Save all databases
            self._save_databases(
                metadata_df, results_df, api_history_df,
                metadata_path, results_path, api_history_path
            )
            
            # Log summary
            self._log_summary(
                processed_images, skipped_images, unique_urls_found, 
                duplicate_urls_skipped, images_folder
            )
            
        except Exception as e:
            logger.error(f"❌ Error in main process: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
    
    def _process_images(self, image_paths: list, metadata_df: pd.DataFrame, 
                       results_df: pd.DataFrame, api_history_df: pd.DataFrame, 
                       current_time_str: str) -> tuple:
        """
        Process all images and perform searches.
        
        Returns:
            tuple: Updated dataframes and processing statistics
        """
        # Track new results for this run
        new_results = []

        # Count skipped and processed images
        skipped_images = 0
        processed_images = 0
        
        # Track unique and duplicate URLs for summary
        unique_urls_found = 0
        duplicate_urls_skipped = 0
        
        for img_path in image_paths:
            img_filename = img_path.name
            
            # Check if image exists in metadata
            img_metadata = metadata_df[metadata_df['filename'] == img_filename]
            
            if not img_metadata.empty:
                # Check if image was processed recently
                last_processed = img_metadata.iloc[0]['last_processed_date']
                linkedin_count = img_metadata.iloc[0].get('linkedin_count', 0)
                
                # Log the processing threshold being used
                if linkedin_count >= self.config["linkedin_high_threshold"]:
                    logger.info(f"📄 Image {img_filename} has {linkedin_count} LinkedIn links (≥{self.config['linkedin_high_threshold']})")
                    logger.info(f"📄 Using {self.config['high_linkedin_days']} day threshold instead of standard {self.config['standard_days']} days")
                
                if self.data_processor.is_recently_processed(
                    last_processed, linkedin_count,
                    self.config["linkedin_high_threshold"],
                    self.config["high_linkedin_days"],
                    self.config["standard_days"]
                ):
                    skipped_images += 1
                    if linkedin_count >= self.config["linkedin_high_threshold"]:
                        logger.info(f"⏭️  Skipped high-LinkedIn image {img_filename} processed within the last {self.config['high_linkedin_days']} days")
                    elif skipped_images % 10 == 0:
                        logger.info(f"⏭️  Skipped {skipped_images} images processed within their threshold periods")
                    continue
                
                # Use existing Imgur URL
                img_url = img_metadata.iloc[0]['imgur_url']
                logger.info(f"♻️  Processing: {img_filename} (using existing Imgur URL)")
                
                # Update last processed date
                metadata_df.loc[metadata_df['filename'] == img_filename, 'last_processed_date'] = current_time_str
            else:
                # New image, upload to Imgur
                logger.info(f"🚀 Processing: {img_filename} (new upload to Imgur)")
                img_url = self.imgur_client.upload_image(img_path)
                
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

            processed_images += 1
            
            # Process searches for this image
            image_results = self._process_image_searches(
                img_filename, img_url, results_df, api_history_df, current_time_str
            )
            
            # Unpack image results
            (image_new_results, api_history_df, 
             image_unique_urls, image_duplicate_urls) = image_results
            
            # Update tracking variables
            new_results.extend(image_new_results)
            unique_urls_found += image_unique_urls
            duplicate_urls_skipped += image_duplicate_urls

            # Log summary for this image
            if image_duplicate_urls > 0:
                logger.info(f"⏩ {img_filename}: Added {image_unique_urls} new URLs, skipped {image_duplicate_urls} existing URLs")
            else:
                logger.info(f"✅ {img_filename}: Added {image_unique_urls} new URLs")
        
        # Add new results to the database
        if new_results:
            logger.info(f"📊 Adding {len(new_results)} new results to database")
            new_results_df = pd.DataFrame(new_results)
            
            if results_df.empty:
                results_df = new_results_df
            else:
                results_df = pd.concat([results_df, new_results_df], ignore_index=True)
        else:
            logger.info("📄 No new results to add to database")
        
        # Update metadata counts and mark duplicates
        logger.info("🔢 Updating metadata link counts")
        metadata_df = self.data_processor.update_metadata_counts(metadata_df, results_df)

        logger.info("🔍 Checking for duplicate URLs in results database")
        results_df = self.data_processor.mark_duplicate_urls(results_df)
        
        return (metadata_df, results_df, api_history_df, new_results,
                skipped_images, processed_images, unique_urls_found, duplicate_urls_skipped)
    
    def _process_image_searches(self, img_filename: str, img_url: str, 
                               results_df: pd.DataFrame, api_history_df: pd.DataFrame,
                               current_time_str: str) -> tuple:
        """
        Process searches for a single image.
        
        Returns:
            tuple: (new_results, updated_api_history_df, unique_urls_count, duplicate_urls_count)
        """
        # Get existing URLs for this image to avoid duplicates
        if not results_df.empty:
            existing_urls = set(results_df[results_df['local_image'] == img_filename]['found_link'].tolist())
        else:
            existing_urls = set()
        
        # Initialize counters for this image
        new_results = []
        unique_urls_count = 0
        duplicate_urls_count = 0
        
        # Process Exact Matches if enabled
        if self.config["do_exact_search"]:
            logger.info(f"🔍 Performing exact matches search for {img_filename}")
            exact_data, api_history_df = self.search_client.search_google_lens(
                img_url, img_filename, api_history_df, "exact_matches"
            )
            
            if exact_data:
                exact_results, exact_new, exact_skipped = self.data_processor.process_search_results(
                    exact_data, "exact_matches", img_filename, img_url, current_time_str, existing_urls
                )
                new_results.extend(exact_results)
                unique_urls_count += exact_new
                duplicate_urls_count += exact_skipped
                # Update existing URLs set for similar search
                existing_urls.update([r["found_link"] for r in exact_results])
            else:
                logger.warning(f"⚠️ No exact matches data received for {img_filename}")
        
        # Process Similar/Visual Matches if enabled
        if self.config["do_similar_search"]:
            logger.info(f"🔍 Performing visual matches search for {img_filename}")
            visual_data, api_history_df = self.search_client.search_google_lens(
                img_url, img_filename, api_history_df, "visual_matches"
            )
            
            if visual_data:
                visual_results, visual_new, visual_skipped = self.data_processor.process_search_results(
                    visual_data, "visual_matches", img_filename, img_url, current_time_str, existing_urls
                )
                new_results.extend(visual_results)
                unique_urls_count += visual_new
                duplicate_urls_count += visual_skipped
            else:
                logger.warning(f"⚠️ No visual matches data received for {img_filename}")
        
        return new_results, api_history_df, unique_urls_count, duplicate_urls_count
    
    def _save_databases(self, metadata_df: pd.DataFrame, results_df: pd.DataFrame,
                       api_history_df: pd.DataFrame, metadata_path: Path,
                       results_path: Path, api_history_path: Path):
        """Save all databases to Excel files."""
        # Save metadata
        logger.info("💾 Saving updated metadata with link counts")
        self.storage_manager.save_excel_with_retry(metadata_df, metadata_path)
        
        # Save results database
        logger.info("💾 Saving updated results database")
        self.storage_manager.save_excel_with_retry(results_df, results_path)
        
        # Save API history database if updated
        logger.info("💾 Saving updated API history database")
        self.storage_manager.save_excel_with_retry(api_history_df, api_history_path)
    
    def _log_summary(self, processed_images: int, skipped_images: int,
                    unique_urls_found: int, duplicate_urls_skipped: int,
                    images_folder: Path):
        """Log final summary statistics."""
        logger.info(f"📊 SUMMARY: Processed {processed_images} images, skipped {skipped_images} images")
        logger.info(f"📊 SUMMARY: Found {unique_urls_found} new URLs, skipped {duplicate_urls_skipped} existing URLs")
        logger.info(f"📊 SUMMARY: API history database updated with new API calls")
        logger.info(f"📊 SUMMARY: Metadata updated with link count statistics")
        logger.info(f"📊 SUMMARY: Using two processing thresholds - {self.config['high_linkedin_days']} days for high-LinkedIn images (≥{self.config['linkedin_high_threshold']}), {self.config['standard_days']} days for others")
        logger.info(f"✅ Process completed.")
        logger.info(f"💾 Excel files saved in: {self.config['metadata_folder']}")
        logger.info(f"🖼️  Images folder: {images_folder}")


def main():
    """Main entry point."""
    logger.info("🚀 Starting Enhanced Image Search Process with API History Tracking")
    logger.info("📊 Custom columns and column widths will be preserved in Excel files")
    
    # Parse arguments and get configuration
    args = parse_arguments()
    config = get_app_config(args)
    
    logger.info(f"📁 Configuration loaded from: {args.config if args.config else 'linkedin_search_image.json'}")
    
    # Initialize and run the application
    app = LinkedInImageSearchApp(config)
    app.run()


if __name__ == "__main__":
    main()


"""
Example usage:

1. Using defaults from linkedin_search_image.json:
python main.py

2. Using a different configuration file:
python main.py --config "path/to/custom_config.json"

3. Overriding specific settings from command line:
python main.py --images-folder "C:\\custom\\images" --standard-days 20

4. Enabling similar matches search (overriding config):
python main.py --enable-similar-search

5. Full example with multiple overrides:
python main.py \
    --config "custom_config.json" \
    --images-folder "C:\\Users\\rober\\custom_images" \
    --linkedin-high-threshold 30 \
    --enable-similar-search

Note: Command line arguments always take precedence over configuration file values.
""" 