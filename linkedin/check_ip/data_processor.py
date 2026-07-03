#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Data Processor Module

This module handles data processing utilities including date extraction, source identification,
and duplicate detection for the LinkedIn image search application.

Author: Roberto
Date: March 2025
"""

import datetime
from datetime import timezone
import logging
import numpy as np
import pandas as pd
import re
from typing import Optional

logger = logging.getLogger(__name__)


class DataProcessor:
    """Handles data processing operations for search results."""
    
    @staticmethod
    def extract_post_date(url: str) -> Optional[str]:
        """
        Extract post date from Twitter/X & LinkedIn URLs.
        
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

    @staticmethod
    def identify_social_media_source(url: str) -> Optional[str]:
        """
        Identify the social media platform from a URL.
        
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

    @staticmethod
    def is_recently_processed(last_processed_date: str, linkedin_count: int, 
                            linkedin_high_threshold: int, high_linkedin_days: int, 
                            standard_days: int) -> bool:
        """
        Check if the image was processed recently based on LinkedIn count.
        
        Args:
            last_processed_date (str): Date string when the image was last processed
            linkedin_count (int): Number of LinkedIn links found for this image
            linkedin_high_threshold (int): Threshold for high LinkedIn count
            high_linkedin_days (int): Days threshold for high LinkedIn count images
            standard_days (int): Days threshold for standard images
            
        Returns:
            bool: True if image was processed recently, False otherwise
        """
        try:
            # Determine the appropriate threshold based on LinkedIn count
            if linkedin_count >= linkedin_high_threshold:
                threshold_days = high_linkedin_days
                logger.debug(f"Using HIGH LinkedIn threshold ({threshold_days} days) for image with {linkedin_count} LinkedIn links")
            else:
                threshold_days = standard_days
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

    @staticmethod
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

    @staticmethod
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

    @staticmethod
    def collect_valid_image_paths(images_folder, allowed_extensions):
        """
        Collect valid image paths from the images folder.
        
        Args:
            images_folder (Path): Path to images folder
            allowed_extensions (set): Set of allowed file extensions
            
        Returns:
            list: List of valid image paths
        """
        return sorted([p for p in images_folder.glob("*.*") if p.suffix.lower() in allowed_extensions])

    @staticmethod
    def process_search_results(search_data: dict, search_type: str, img_filename: str, 
                             img_url: str, current_time_str: str, existing_urls: set) -> tuple:
        """
        Process search results and extract relevant information.
        
        Args:
            search_data (dict): Search response data
            search_type (str): Type of search performed
            img_filename (str): Image filename
            img_url (str): Imgur URL
            current_time_str (str): Current timestamp string
            existing_urls (set): Set of existing URLs to avoid duplicates
            
        Returns:
            tuple: (new_results_list, new_urls_count, skipped_urls_count)
        """
        new_results = []
        new_urls_count = 0
        skipped_urls_count = 0
        
        # Determine the key to use based on search type
        if search_type == "exact_matches":
            matches_key = "exact_matches"
            match_type = "Exact Match"
        elif search_type == "visual_matches":
            matches_key = "visual_matches"
            match_type = "Similar Match"
        else:
            return new_results, new_urls_count, skipped_urls_count
        
        matches = search_data.get(matches_key, [])
        
        for index, res in enumerate(matches, start=1):
            link = res.get("link", "")
            
            # Skip if this URL already exists for this image
            if link in existing_urls:
                skipped_urls_count += 1
                continue
            
            new_urls_count += 1
            
            title = res.get("title", "")
            post_date = DataProcessor.extract_post_date(link)
            source = DataProcessor.identify_social_media_source(link)
            
            new_results.append({
                "local_image": img_filename,
                "uploaded_url": img_url,
                "found_link": link,
                "title": title,
                "duplicate": 0,  # Initialize as non-duplicate
                "match_type": match_type,
                "source": source,
                "post_date": post_date,
                "search_date": current_time_str,
                "order": index  # Add order of appearance
            })
        
        return new_results, new_urls_count, skipped_urls_count 