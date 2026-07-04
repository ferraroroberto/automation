#!/usr/bin/env python3
"""
Illustrations Subscriber Processing Script

This script processes illustrations and distributes them to subscriber folders
based on metadata from Excel files. It can also encode hidden messages into
the illustrations using LSB steganography.

Features:
- Filters subscribers based on subscription type and update date
- Filters illustrations based on tags and social media metrics
- Copies illustrations to subscriber folders
- Encodes ownership messages into images using steganography
- Configurable via JSON or command-line arguments

Usage:
    python illustrations_subscribers.py [options]

Options:
    --config PATH      Path to config file (default: illustrations_subscribers.json)
    --subscribers PATH Path to subscribers Excel file
    --illustrations PATH Path to illustrations Excel file
    --source PATH      Path to source illustrations folder
    --archive PATH     Path to archive folder
    --encode {y,n}     Whether to encode messages (default: prompt user)
    --log-level LEVEL  Logging level (default: INFO)
"""

import argparse
import glob
import json
import logging
import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

import numpy as np
import pandas as pd
from PIL import Image

# Set up logging
def setup_logging(log_level: str = "INFO") -> logging.Logger:
    """Configure logging with the specified log level."""
    numeric_level = getattr(logging, log_level.upper(), None)
    if not isinstance(numeric_level, int):
        raise ValueError(f"Invalid log level: {log_level}")
    
    logging.basicConfig(
        level=numeric_level,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    return logging.getLogger(__name__)

# Function to encode a message into an image
def encode_message(image_path: str, message: str, output_path: str) -> None:
    """
    Encodes a hidden message into an image using LSB steganography.

    Uses numpy array slicing + bit-shift over a flat pixel buffer instead of
    per-pixel getpixel()/putpixel() calls, which is O(w*h) in pure Python
    (audit issue #67).

    Parameters:
    - image_path: Path to the input image.
    - message: The message to encode into the image.
    - output_path: Path to save the output image with the encoded message.
    """
    img = Image.open(image_path)
    pixels = np.array(img)
    if pixels.ndim != 3:
        raise ValueError(f"encode_message requires a color image, got mode {img.mode!r}")

    # Convert the message into a binary string
    binary_message = ''.join(format(ord(char), '08b') for char in message)
    binary_message += '1111111111111110'  # Delimiter to indicate the end of the message
    bits = np.frombuffer(binary_message.encode('ascii'), dtype=np.uint8) - ord('0')

    # Only the RGB planes are modified (any alpha/extra channel is left untouched),
    # matching the original per-pixel behavior.
    n_color_channels = min(3, pixels.shape[-1])
    rgb = pixels[..., :n_color_channels].copy()
    flat = rgb.reshape(-1)
    if bits.size > flat.size:
        raise ValueError("Message is too long to encode in this image")

    # Clear the LSB of each targeted byte, then OR in the message bit
    flat[:bits.size] = (flat[:bits.size] & 0xFE) | bits
    pixels[..., :n_color_channels] = flat.reshape(rgb.shape)

    # Save the image with the encoded message
    encoded_img = Image.fromarray(pixels, mode=img.mode)
    encoded_img.save(output_path)

def load_dataframe(file_path: str, logger: logging.Logger, description: str = "file", max_retries: int = 3) -> pd.DataFrame:
    """
    Load a DataFrame from an Excel file with error handling.

    Retries only on the "file locked/missing" error family (FileNotFoundError,
    PermissionError) and only up to max_retries times, so a persistently
    missing/locked file raises instead of looping forever (audit issue #67).
    """
    attempt = 0
    while True:
        try:
            df = pd.read_excel(file_path)
            logger.info(f"Successfully loaded {description} from {file_path}")
            return df
        except FileNotFoundError:
            attempt += 1
            if attempt >= max_retries:
                logger.error(f"Error: {description.capitalize()} not found at {file_path} after {max_retries} attempts")
                raise
            logger.error(f"Error: {description.capitalize()} not found at {file_path}")
            input("Press any key to try again...")
        except PermissionError:
            attempt += 1
            if attempt >= max_retries:
                logger.error(f"Error: {description.capitalize()} is open at {file_path} after {max_retries} attempts")
                raise
            logger.error(f"Error: {description.capitalize()} is open. Please close it and try again.")
            input("Press any key to try again...")

def parse_arguments() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(description="Process illustrations for subscribers")
    parser.add_argument("--config", type=str, 
                        default=os.path.splitext(os.path.basename(__file__))[0] + ".json",
                        help="Path to config file")
    parser.add_argument("--subscribers", type=str, help="Path to subscribers Excel file")
    parser.add_argument("--illustrations", type=str, help="Path to illustrations Excel file")
    parser.add_argument("--source", type=str, help="Path to source illustrations folder")
    parser.add_argument("--archive", type=str, help="Path to archive folder")
    parser.add_argument("--encode", choices=["y", "n"], help="Whether to encode messages")
    parser.add_argument("--log-level", type=str, default="INFO", 
                        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
                        help="Logging level")
    return parser.parse_args()

def load_config(config_path: str, args: argparse.Namespace, logger: logging.Logger) -> Dict[str, Any]:
    """Load configuration from JSON file and override with command-line arguments."""
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
        logger.info(f"Configuration loaded from {config_path}")
    except Exception as e:
        logger.error(f"Error loading config file: {str(e)}. Using default values.")
        config = {
            "metadata_subscribers": "",
            "metadata_archive_folder": "",
            "metadata_illustrations": "",
            "metadata_source_folder": "",
            "log_level": "INFO"
        }

    # Override config with command-line arguments if provided
    if args.subscribers:
        config["metadata_subscribers"] = args.subscribers
    if args.illustrations:
        config["metadata_illustrations"] = args.illustrations
    if args.source:
        config["metadata_source_folder"] = args.source
    if args.archive:
        config["metadata_archive_folder"] = args.archive
    if args.log_level:
        config["log_level"] = args.log_level
    if args.encode:
        config["encode"] = args.encode

    # Convert paths to use proper OS path separators
    for key in ["metadata_subscribers", "metadata_archive_folder", 
                "metadata_illustrations", "metadata_source_folder"]:
        if key in config:
            config[key] = os.path.normpath(config[key])

    return config

def process_subscribers(config: Dict[str, Any], logger: logging.Logger) -> None:
    """Main processing function."""
    # Get today's date in the appropriate format
    today_date = datetime.now().date()
    
    # Load subscriber metadata
    subscribers_df = load_dataframe(config["metadata_subscribers"], logger, "subscriber metadata")
    
    # Load illustration metadata
    illustrations_df = load_dataframe(config["metadata_illustrations"], logger, "illustration metadata")
    
    # Filter subscribers based on the 'sub type' and 'update' columns
    filtered_subscribers_list = []

    for index, row in subscribers_df.iterrows():
        # Check if the 'sub type' is 'Wanted illustrations - active'
        if row['sub type'] in ['Wanted illustrations - active', 'Wanted illustrations - gift']:
            # Check if the 'update' column exists and is not empty
            if 'update' in row and pd.notnull(row['update']):
                # Check if the 'update' column matches today's date
                if pd.to_datetime(row['update']).date() == today_date:
                    filtered_subscribers_list.append(row)  # Append the row to the list
                    logger.info(f"Processing record for {row['name']} as it has today's update date and is an active wanted illustration.")
                else:
                    logger.debug(f"Skipping record for {row['name']} due to unmatched update date.")
            else:
                logger.warning(f"The 'update' field is missing or empty for {row['name']}. Please ensure the field is properly filled.")

    # Convert the list of rows to a DataFrame
    filtered_subscribers_df = pd.DataFrame(filtered_subscribers_list)

    if filtered_subscribers_df.empty:
        logger.warning("No subscribers found with today's update date and 'Wanted illustrations - active' type. Please check the subscriber metadata.")
        return
    
    # Continue processing if there are valid subscribers
    for name in filtered_subscribers_df['name']:
        # Create or clean subscriber folders
        subscriber_folder = os.path.join(config["metadata_archive_folder"], name)
        os.makedirs(subscriber_folder, exist_ok=True)  # Create the folder if it does not exist
        
        # Check if there are any subfolders
        subfolders = [f.path for f in os.scandir(subscriber_folder) if f.is_dir()]
        if subfolders:
            user_input = input(f"Do you want to delete the existing subfolders in {subscriber_folder}? (y/n): ").strip().lower()
            if user_input == 'y':
                shutil.rmtree(subscriber_folder)  # Remove existing folder and its contents
                os.makedirs(subscriber_folder)  # Create a new clean folder
                logger.info(f"Cleaned existing folder: {subscriber_folder}")
            else:
                logger.info(f"Skipping deletion of subfolders in {subscriber_folder}.")
        else:
            logger.debug(f"No subfolders found in {subscriber_folder}.")

    # Filter illustrations based on tags and social media metrics
    filtered_illustrations_df = illustrations_df[
        (illustrations_df['FK_TAGS'] == 'archived') &
        ((illustrations_df['N_TIMES_IG'] +
          illustrations_df['N_TIMES_LI'] +
          illustrations_df['N_TIMES_TW'] +
          illustrations_df['N_TIMES_CMT'] +
          illustrations_df['N_TIMES_PST']) > 0)
    ]

    # Initialize counters for logging
    total_copied = 0
    total_skipped = 0
    total_files_processed = 0
    total_unique_illustrations = 0

    # Copy illustrations to subscriber folders
    skipped_files = []

    for de_illustration in filtered_illustrations_df['DE_ILLUSTRATION']:
        source_file = os.path.join(config["metadata_source_folder"], f"{de_illustration}.png")

        if os.path.exists(source_file):
            # Copy the file to each subscriber's folder
            for name in filtered_subscribers_df['name']:
                destination_file = os.path.join(config["metadata_archive_folder"], name, f"{de_illustration}.png")
                shutil.copy2(source_file, destination_file)
            total_copied += 1
            total_files_processed += 1
        else:
            # Log skipped files
            skipped_files.append(de_illustration)
            total_skipped += 1

    # Attempt to find similar named files for skipped files
    for skipped_file in skipped_files:
        pattern = os.path.join(config["metadata_source_folder"], f"{skipped_file}*.png")
        matched_files = glob.glob(pattern)

        if matched_files:
            # Copy any similar files found
            for matched_file in matched_files:
                for name in filtered_subscribers_df['name']:
                    destination_file = os.path.join(config["metadata_archive_folder"], name, os.path.basename(matched_file))
                    shutil.copy2(matched_file, destination_file)
            total_files_processed += len(matched_files)
            total_unique_illustrations += 1
        else:
            # Log if no matching files are found
            logger.warning(f"No matching files found for: {skipped_file}")

    # Final report
    logger.info("Final process report:")
    logger.info(f"Total files processed: {total_files_processed}")
    logger.info(f"Total unique illustrations records: {total_copied + total_unique_illustrations}")
    logger.info(f"Total skipped files with no matching patterns: {total_skipped - total_unique_illustrations}")

    # Determine whether to encode messages
    encode_option = config.get("encode", None)
    if encode_option is None:
        user_input = input("Do you want to encode the message into the files? (Y/N): ")
        encode_option = user_input.lower()
    
    if encode_option.lower() == 'y':
        log_each_file = input("Do you want to log each encoded file? (Y/N): ")
        start_time = time.time()
        encoding_counter = 0

        # Encode message into each file in subscriber folders
        for name in filtered_subscribers_df['name']:
            subscriber_folder = os.path.join(config["metadata_archive_folder"], name)
            for file_name in os.listdir(subscriber_folder):
                if file_name.endswith('.png'):
                    file_path = os.path.join(subscriber_folder, file_name)
                    message = f"copy for {name}. For personal use only, according to guidelines at https://robertoferraro.art/. No copy or further distribution is allowed"
                    encode_message(file_path, message, file_path)
                    encoding_counter += 1
                    if log_each_file.lower() == 'y':
                        logger.info(f"Encoded file: {file_path}")
                    if encoding_counter % 50 == 0:
                        logger.info(f"Encoded {encoding_counter} files...")

        end_time = time.time()
        encoding_time = end_time - start_time
        logger.info(f"Total files encoded: {encoding_counter}")
        logger.info(f"Total time for encoding: {encoding_time:.2f} seconds")
    else:
        logger.info("Encoding process skipped.")

def main() -> None:
    """Main entry point for the script."""
    args = parse_arguments()
    
    # Setup logging with default level
    logger = setup_logging(args.log_level)
    
    # Determine config file path (absolute or relative to script)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if not os.path.isabs(args.config):
        config_path = os.path.join(script_dir, args.config)
    else:
        config_path = args.config
    
    # Load configuration
    config = load_config(config_path, args, logger)
    
    # Update logging level if specified in config
    if config.get("log_level") != args.log_level:
        logger = setup_logging(config["log_level"])
    
    # Process subscribers
    process_subscribers(config, logger)

if __name__ == "__main__":
    main()