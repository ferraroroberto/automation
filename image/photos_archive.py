"""
Photo Archive Automation Script

Organizes photos and videos with intelligent duplicate detection and metadata management.
Extracts creation dates from EXIF, video metadata, and filename patterns.
Creates unified naming: YYYYMMDD-HHMMSS.extension

Configuration: photos_archive.json (follows AGENTS.md guidelines)
- source_folder: Path to scan for media files
- destination_folder: Path to store organized archive
- processing_thresholds: Time/file count thresholds
- behavior_flags: Control prompts and automation behavior
- file_extensions: Supported image/video formats
- date_parsing: Regex patterns for filename date extraction

References:
- https://chatgpt.com/c/9887f68a-8392-40a9-a648-321dc77ff475 (long version)
- https://chatgpt.com/c/9aaeba93-8e88-416f-a502-0e4d4bfe9270 (video dates)
- https://chatgpt.com/c/c8d6ca75-c05d-4a35-9360-dae22274c732 (pattern recognition)
- https://chatgpt.com/c/d81cead9-9fe0-49f7-a1a8-28d2d45c21d1 (deleting discards)
- https://chatgpt.com/c/a5bf8310-98ab-4626-8e48-d223bb54d7f7 (source/dest same path)
"""

import os
import re
import json
import pandas as pd
import shutil
from datetime import datetime, timedelta
from PIL import Image
from PIL.ExifTags import TAGS
import logging
import hashlib
import tkinter as tk
from tkinter import filedialog
from pymediainfo import MediaInfo

def load_config():
    """Load and validate configuration from photos_archive.json."""
    config_path = os.path.join(os.path.dirname(__file__), 'photos_archive.json')
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

# Load configuration
CONFIG = load_config()

def calculate_time_intervals(df):
    """Calculate time intervals between consecutive files within same day."""
    # Extract date part from unified names (YYYYMMDD)
    df['unified_day'] = df['unified_name'].apply(lambda x: x[:8])

    # Convert to datetime objects
    df['datetime'] = df['unified_name'].apply(lambda x: datetime.strptime(x[:15], '%Y%m%d-%H%M%S'))

    # Calculate seconds between consecutive files per day
    df['time_interval'] = df.groupby('unified_day')['datetime'].diff().dt.total_seconds().fillna(0)

    return df

def calculate_short_sequence_counts(df, threshold=None):
    """Count files in short sequences per day (below time threshold)."""
    if threshold is None:
        threshold = CONFIG['processing_thresholds']['short_sequence_seconds']

    # Ensure time intervals are calculated
    if 'time_interval' not in df.columns:
        df = calculate_time_intervals(df)

    # Mark files with short intervals
    df['short_sequence'] = df['time_interval'] <= threshold
    df['short_sequence_count'] = df.groupby('unified_day')['short_sequence'].transform('sum')

    # Clean up temporary columns
    df.drop(columns=['datetime', 'time_interval', 'short_sequence'], inplace=True)

    return df

def setup_logging(log_folder=None):
    """Configure logging to timestamped file in specified folder."""
    if log_folder is None:
        log_folder = CONFIG.get('log_folder', CONFIG['destination_folder'])

    timestamp = datetime.now().strftime('%Y%m%d-%H%M')
    log_file = os.path.join(log_folder, f'photo_processing_{timestamp}.log')

    os.makedirs(log_folder, exist_ok=True)

    # Set logging level from config
    log_level = getattr(logging, CONFIG['logging']['level'].upper(), logging.INFO)

    logging.basicConfig(
        level=log_level,
        filename=log_file,
        filemode='w',
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    logging.info(f"Logging initialized: {log_file}")

def get_exif_creation_date(file_path):
    """Extract creation date from EXIF DateTimeOriginal tag."""
    try:
        image = Image.open(file_path)
        exif_data = image._getexif()
        if exif_data:
            for tag, value in exif_data.items():
                decoded = TAGS.get(tag, tag)
                if decoded == "DateTimeOriginal":
                    # Try standard EXIF format first
                    try:
                        return datetime.strptime(value, "%Y:%m:%d %H:%M:%S")
                    except ValueError:
                        # Try alternative format
                        try:
                            return datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
                        except ValueError:
                            # Handle invalid hour 24 by rolling to next day
                            if " 24:" in value:
                                corrected_value = value.replace(" 24:", " 00:")
                                corrected_datetime = datetime.strptime(corrected_value, "%Y:%m:%d %H:%M:%S") + timedelta(days=1)
                                return corrected_datetime
                            logging.warning(f"Unknown EXIF date format: {value}")
                            return None
    except Exception as e:
        logging.warning(f"EXIF extraction failed for {file_path}: {e}")
    return None


def parse_date(date_str):
    """
    Parses a date string and returns a datetime object.

    Args:
    date_str (str): Date string to parse.

    Returns:
    datetime: Parsed datetime object, or None if parsing fails.
    """
    try:
        return datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S %Z")
    except ValueError:
        try:
            return datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S.%fZ")
        except ValueError:
            logging.warning(f"Unknown date format for video date: {date_str}")
            return None

def extract_date_from_filename(filename):
    """Extract date from filename using configured regex patterns."""
    patterns = CONFIG['date_parsing']['filename_patterns']
    for pattern in patterns:
        match = re.search(pattern, filename, re.IGNORECASE)
        if match:
            # Extract and normalize date/time components
            date_part = match.group(1).replace('-', '')
            time_part = match.group(2).replace('.', '').replace('_', '').replace('-', '') if len(match.groups()) > 1 else '000000'
            date_str = date_part + time_part
            try:
                return datetime.strptime(date_str, '%Y%m%d%H%M%S')
            except ValueError as ve:
                logging.error(f"Date parsing error: {ve}")
                continue
    return None

def get_video_creation_date(file_path):
    """Extract creation date from video metadata using pymediainfo."""
    try:
        media_info = MediaInfo.parse(file_path)
        for track in media_info.tracks:
            if track.track_type == "General":
                # Try encoded_date, tagged_date, or recorded_date
                creation_date_str = track.encoded_date or track.tagged_date or track.recorded_date
                if creation_date_str:
                    return parse_date(creation_date_str)
        logging.warning(f"No creation date in video metadata: {file_path}")
        return None
    except Exception as e:
        logging.warning(f"Video metadata extraction failed for {file_path}: {e}")
        return None


def get_file_info(file_path):
    """Extract comprehensive metadata from media file."""
    try:
        file_stat = os.stat(file_path)
        creation_time = datetime.fromtimestamp(file_stat.st_ctime)
        modified_time = datetime.fromtimestamp(file_stat.st_mtime)

        # Initialize defaults
        file_type = 'Unknown'
        exif_date = None
        video_date = None
        criteria = "modified"
        unified_name = modified_time.strftime('%Y%m%d-%H%M%S') + os.path.splitext(file_path)[1].lower()

        # Determine file type and extract creation date
        file_extension_lower = file_path.lower()
        image_extensions = [ext.lower() for ext in CONFIG['file_extensions']['image']]
        video_extensions = [ext.lower() for ext in CONFIG['file_extensions']['video']]

        if any(file_extension_lower.endswith(ext) for ext in image_extensions):
            file_type = 'Image'
            exif_date = get_exif_creation_date(file_path)
            if exif_date:
                creation_time = exif_date
                criteria = "exif"
        elif any(file_extension_lower.endswith(ext) for ext in video_extensions):
            file_type = 'Video'
            video_date = get_video_creation_date(file_path)
            if video_date:
                creation_time = video_date
                criteria = "video"

        # If no EXIF or video date is found, try to guess from filename
        if not exif_date and not video_date:
            guessed_date = extract_date_from_filename(os.path.basename(file_path))
            if guessed_date:
                creation_time = guessed_date
                criteria = "filename"
            else:
                # If no guessed date, take the smallest between creation and modified dates
                if creation_time > modified_time:
                    creation_time = modified_time
                    criteria = "modified"
                else:
                    criteria = "creation"

        unified_name = creation_time.strftime('%Y%m%d-%H%M%S') + os.path.splitext(file_path)[1].lower()
        dest_path = os.path.join(CONFIG['destination_folder'], unified_name)

        # Check if the file is in the destination directory (but not in its subfolders)
        if os.path.abspath(os.path.dirname(file_path)) == os.path.abspath(CONFIG['destination_folder']):
            exclude = 1  # Exclude if the file is directly in the destination directory
        else:
            exclude = 0  # Do not exclude if the file is in a subfolder

        # Return file metadata as a dictionary
        return {
            'file_path': file_path,
            'file_name': os.path.basename(file_path),
            'file_extension': os.path.splitext(file_path)[1].lower(),
            'file_type': file_type,
            'file_size': file_stat.st_size,
            'creation_date': creation_time,
            'modified_date': modified_time,
            'exif_date': exif_date,
            'video_date': video_date,
            'criteria': criteria,
            'unified_name': unified_name,
            'destination_path': dest_path,
            'duplicate': 0,
            'duplicate_count': 0,
            'discard': 0,
            'sha256': '',
            'copy_success': 0,
            'deleted': 0,
            'exclude': 0
        }
    except Exception as e:
        # Log error if file information cannot be retrieved
        logging.error(f"Error getting file info for {file_path}: {e}")
        return None

def process_files(source_folder, dest_folder, skip_dest_folder):
    """
    Processes all files in the source folder and its subfolders to extract metadata.
    Optionally excludes files in the destination folder if it's a subfolder of the source folder.

    Args:
    source_folder (str): Path to the source folder.
    dest_folder (str): Path to the destination folder.
    skip_dest_folder (bool): Whether to skip the destination folder during scanning.

    Returns:
    pd.DataFrame: DataFrame containing metadata for all files.
    """
    file_records = []

    # Traverse the directory structure
    for root, dirs, files in os.walk(source_folder):
        # Exclude the destination folder from the file collection process if specified
        if skip_dest_folder and os.path.commonpath([root, dest_folder]) == dest_folder:
            continue

        for file in files:
            file_path = os.path.join(root, file)
            file_info = get_file_info(file_path)
            if file_info:
                file_records.append(file_info)
            # Log and print progress every 1000 files
            if len(file_records) % 1000 == 0:
                logging.info(f"{len(file_records)} files processed")
                print(f"{len(file_records)} files processed")

    return pd.DataFrame(file_records)

def delete_discarded_files(df):
    """
    Deletes the source files that are marked as discarded in the metadata.

    Args:
    df (pd.DataFrame): DataFrame containing file metadata.
    """
    deleted_files_count = 0
    last_logged_count = 0

    for index, row in df.iterrows():
        if row['deleted'] == 0 and row['discard'] == 1:
            try:
                os.remove(row['file_path'])
                df.at[index, 'deleted'] = 1
                deleted_files_count += 1
            except FileNotFoundError:
                logging.warning(f"File not found for deletion: {row['file_path']}")
                df.at[index, 'deleted'] = -1
            except Exception as e:
                logging.error(f"Error deleting file {row['file_path']}: {e}")
                df.at[index, 'deleted'] = -1

        if deleted_files_count // 1000 > last_logged_count // 1000:
            logging.info(f"{deleted_files_count} discarded files deleted")
            print(f"{deleted_files_count} discarded files deleted")
            last_logged_count = deleted_files_count

    logging.info(f"{deleted_files_count} discarded files deleted")
    print(f"{deleted_files_count} discarded files deleted")

def calculate_sha256(file_path):
    """
    Calculate the SHA256 hash of a file.

    Args:
    file_path (str): Path to the file.

    Returns:
    str: SHA256 hash of the file.
    """
    sha256_hash = hashlib.sha256()
    with open(file_path, 'rb') as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def mark_duplicates(df):
    """
    Identifies and marks duplicate files in the DataFrame and counts duplicates.

    Args:
    df (pd.DataFrame): DataFrame containing file metadata.

    Returns:
    pd.DataFrame: DataFrame with duplicates marked in the 'duplicate', 'duplicate_count', and 'discard' columns.
    """
    df['duplicate'] = df.duplicated(subset=['unified_name'], keep=False).astype(int)
    df['duplicate_count'] = df.groupby('unified_name')['unified_name'].transform('count')
    df['discard'] = 0

    # Identify groups of potential duplicates
    potential_duplicates = df[df['duplicate'] == 1]

    # Count the number of duplicates
    num_duplicates = len(potential_duplicates)
    if num_duplicates > 0:
        check_sha256 = CONFIG['duplicate_detection']['enable_sha256_check']
        if not check_sha256:
            check_sha256_input = input(f"Do you want to check SHA256 for the {num_duplicates} duplicates? (Y/N): ").strip().lower()
            check_sha256 = check_sha256_input == 'y'
        if check_sha256:
            sha256_checked_count = 0
            for index, row in potential_duplicates.iterrows():
                sha256 = calculate_sha256(row['file_path'])
                df.at[index, 'sha256'] = sha256
                sha256_checked_count += 1

                # Log and print progress every 50 files
                if sha256_checked_count % 50 == 0:
                    logging.info(f"{sha256_checked_count} files checked for SHA256")
                    print(f"{sha256_checked_count} files checked for SHA256")

            # Update discard column based on SHA256
            for name in df['unified_name'].unique():
                subset = df[df['unified_name'] == name]
                if len(subset) > 1:
                    subset = subset.dropna(subset=['sha256'])  # Remove entries without SHA256 hash
                    subset_sorted = subset.sort_values(by=['sha256', 'file_size'], ascending=[True, False])
                    sha256_groups = subset_sorted.groupby('sha256')
                    for sha256, group in sha256_groups:
                        df.loc[group.index[1:], 'discard'] = 1

    return df

def calculate_total_files_unified_day(df):
    """
    Calculates the total number of files with the same unified name in year, month, and day.

    Args:
    df (pd.DataFrame): DataFrame containing file metadata.

    Returns:
    pd.DataFrame: DataFrame with the total_files_unified_day column added.
    """
    df['unified_day'] = df['unified_name'].apply(lambda x: x[:8])  # Extract YYYYMMDD from the unified name
    df['total_files_unified_day'] = df.groupby('unified_day')['unified_day'].transform('count')
    return df

def recalculate_metadata(df):
    """
    Recalculates the metadata fields based on the existing file information.

    Args:
    df (pd.DataFrame): DataFrame containing file metadata.

    Returns:
    pd.DataFrame: DataFrame with recalculated metadata fields.
    """
    def generate_unified_name(row):
        # Handle NaT (Not a Time) cases by checking if dates are valid
        if pd.isna(row['creation_date']):
            date_to_use = row['modified_date']
        else:
            date_to_use = row['creation_date']

        # If both dates are NaT or invalid, set a default date (e.g., the epoch start date)
        if pd.isna(date_to_use):
            date_to_use = datetime(1970, 1, 1)

        # Handle the case where the file extension might be NaN
        file_extension = row['file_extension'] if isinstance(row['file_extension'], str) else '.unknown'

        return date_to_use.strftime('%Y%m%d-%H%M%S') + file_extension

    # Recalculate the unified_name field
    df['unified_name'] = df.apply(generate_unified_name, axis=1)
    df['duplicate'] = df.duplicated(subset=['unified_name'], keep=False).astype(int)
    df['duplicate_count'] = df.groupby('unified_name')['unified_name'].transform('count')
    df['discard'] = 0
    df['copy_success'] = 0
    df['deleted'] = 0

    # Apply the exclusion logic again
    df['exclude'] = df.apply(
        lambda row: 1 if os.path.abspath(os.path.dirname(row['file_path'])) == os.path.abspath(
            CONFIG['destination_folder']) else 0,
        axis=1
    )

    # Calculate total_files_unified_day and time intervals
    df = calculate_total_files_unified_day(df)
    df = calculate_time_intervals(df)
    df = calculate_short_sequence_counts(df)

    # Mark duplicates with the same file name and unified name for discard
    for name in df['unified_name'].unique():
        subset = df[df['unified_name'] == name]
        if len(subset) > 1:
            df.loc[subset.index, 'duplicate'] = 1
            for file_name in subset['file_name'].unique():
                sub_subset = subset[subset['file_name'] == file_name]
                if len(sub_subset) > 1:
                    sub_subset_sorted = sub_subset.sort_values(by='file_size', ascending=False)
                    df.loc[sub_subset_sorted.index[1:], 'discard'] = 1

    # Recalculate destination_path
    df['destination_path'] = df['unified_name'].apply(lambda x: os.path.join(CONFIG['destination_folder'], x))

    # Display top unified days only once
    top_days = df['unified_day'].value_counts().head(20)
    for idx, (day, count) in enumerate(top_days.items(), start=1):
        log_msg = f"#{idx:02d} - {day} - {count} files"
        print(log_msg)
        logging.info(log_msg)

    return df

def copy_files(df):
    """
    Copies non-duplicate files to the destination folder, renaming them based on the unified name.

    Args:
    df (pd.DataFrame): DataFrame containing file metadata.
    """
    copied_files_count = 0
    last_logged_count = 0
    name_count = {}

    for index, row in df.iterrows():
        if row['discard'] == 0 and row['copy_success'] == 0 and row['exclude'] == 0:
            if row['duplicate_count'] > 1:
                # Add suffix if the unified name is duplicated
                if row['unified_name'] in name_count:
                    name_count[row['unified_name']] += 1
                else:
                    name_count[row['unified_name']] = 1

                suffix = f"-{name_count[row['unified_name']]:03d}"
                unified_name_base = row['unified_name'].rsplit('.', 1)[0]  # Remove the extension from unified name
                new_file_name = f"{unified_name_base}{suffix}{row['file_extension']}"
            else:
                new_file_name = row['unified_name']  # No suffix for non-duplicates

            dest_path = os.path.join(CONFIG['destination_folder'], new_file_name)
            df.at[index, 'destination_path'] = dest_path

            # Skip files already present in the destination folder
            if os.path.exists(dest_path):
                df.at[index, 'copy_success'] = 1  # Assume already copied successfully
                continue

            try:
                shutil.copy2(row['file_path'], dest_path)
                df.at[index, 'copy_success'] = 1
                copied_files_count += 1
            except Exception as e:
                logging.error(f"Error copying file {row['file_path']} to {dest_path}: {e}")

        # Log and print progress every 1000 files, only once for each milestone
        if copied_files_count // 1000 > last_logged_count // 1000:
            logging.info(f"{copied_files_count} files copied")
            print(f"{copied_files_count} files copied")
            last_logged_count = copied_files_count

    logging.info(f"{copied_files_count} files copied")
    print(f"{copied_files_count} files copied")

def delete_copied_files(df):
    """
    Deletes the source files that were successfully copied to the destination folder.

    Args:
    df (pd.DataFrame): DataFrame containing file metadata.
    """
    deleted_files_count = 0
    last_logged_count = 0

    for index, row in df.iterrows():
        if row['copy_success'] == 1:
            try:
                os.remove(row['file_path'])
                df.at[index, 'deleted'] = 1
                deleted_files_count += 1
            except Exception as e:
                logging.error(f"Error deleting file {row['file_path']}: {e}")
                df.at[index, 'deleted'] = -1

        # Log and print progress every 1000 files, only once for each milestone
        if deleted_files_count // 1000 > last_logged_count // 1000:
            logging.info(f"{deleted_files_count} files deleted")
            print(f"{deleted_files_count} files deleted")
            last_logged_count = deleted_files_count

    logging.info(f"{deleted_files_count} files deleted")
    print(f"{deleted_files_count} files deleted")

# Call this function at the beginning of your main function
def main(source_folder, dest_folder):
    # Set up logging to save log file in the destination folder
    current_time_str = datetime.now().strftime('%Y%m%d-%H%M')
    setup_logging(dest_folder)

    # Check if user wants to use existing metadata
    use_existing_metadata = CONFIG['behavior_flags']['use_existing_metadata']
    if not use_existing_metadata:
        use_existing_metadata_input = input("Do you want to use an existing metadata file? (Y/N): ").strip().lower()
        use_existing_metadata = use_existing_metadata_input == 'y'
    if use_existing_metadata == 'y':
        # Ensure Tkinter is properly initialized
        root = tk.Tk()
        root.withdraw()  # Hide the root window

        # Get the screen width and height
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # Set the geometry of the root window to the center of the screen
        window_width = 300  # Width of the dialog window
        window_height = 200  # Height of the dialog window
        position_right = int(screen_width/2 - window_width/2)
        position_down = int(screen_height/2 - window_height/2)

        # Adjust the geometry
        root.geometry(f"{window_width}x{window_height}+{position_right}+{position_down}")

        # Open the file dialog
        metadata_file = filedialog.askopenfilename(title="Select metadata Excel file", filetypes=[("Excel files", "*.xlsx")])
        root.destroy()  # Properly destroy the Tkinter root window
        if metadata_file:
            df = pd.read_excel(metadata_file)
            logging.info(f"Loaded metadata from {metadata_file}")
            print(f"Loaded metadata from {metadata_file}")

            # Check if user wants to recalculate the logic
            recalculate_logic = CONFIG['behavior_flags']['recalculate_logic']
            if not recalculate_logic:
                recalculate_logic_input = input("Do you want to recalculate the logic? (Y/N): ").strip().lower()
                recalculate_logic = recalculate_logic_input == 'y'
            if recalculate_logic:
                df = recalculate_metadata(df)

                # Display top unified days
                top_days = df['unified_day'].value_counts().head(20)
                for idx, (day, count) in enumerate(top_days.items(), start=1):
                    log_msg = f"#{idx:02d} - {day} - {count} files"
                    print(log_msg)
                    logging.info(log_msg)

                # Ask for exclusion threshold
                default_exclude_threshold = CONFIG['processing_thresholds']['exclude_days_over_files']
                exclude_threshold_input = input(f"Do you want to exclude days with more than X files? Enter X value, or zero if you don't want to exclude any file (default: {default_exclude_threshold}): ").strip()
                exclude_threshold = int(exclude_threshold_input) if exclude_threshold_input else default_exclude_threshold
                if exclude_threshold > 0:
                    df.loc[df['total_files_unified_day'] > exclude_threshold, 'exclude'] = 1

                # Ask for short sequence exclusion threshold
                default_short_seq_threshold = CONFIG['processing_thresholds']['exclude_short_sequences_over']
                short_seq_threshold_input = input(f"Enter the maximum number of files in short sequence to exclude days (default: {default_short_seq_threshold}): ").strip()
                short_seq_threshold = int(short_seq_threshold_input) if short_seq_threshold_input else default_short_seq_threshold
                if short_seq_threshold > 0:
                    df.loc[df['short_sequence_count'] > short_seq_threshold, 'exclude'] = 1

                # Check to exclude files based on criteria
                exclude_criteria = CONFIG['behavior_flags']['exclude_creation_modified_criteria']
                if not exclude_criteria:
                    exclude_criteria_input = input("Do you want to exclude files where the unified name is based on creation or modified dates? (Y/N): ").strip().lower()
                    exclude_criteria = exclude_criteria_input == 'y'
                if exclude_criteria:
                    df.loc[df['criteria'].isin(['creation', 'modified']), 'exclude'] = 1

                current_time_str = datetime.now().strftime('%Y%m%d-%H%M')
                metadata_file = os.path.join(dest_folder, f'metadata_{current_time_str}.xlsx')
                df.to_excel(metadata_file, index=False)
                logging.info(f"Recalculated metadata and saved to {metadata_file}")
                print(f"Recalculated metadata and saved to {metadata_file}")

            # Check for confirmation to process files
            continue_copy = CONFIG['behavior_flags']['continue_copy']
            if not continue_copy:
                continue_copy_input = input("Do you want to continue with copying the files? (Y/N): ").strip().lower()
                continue_copy = continue_copy_input == 'y'
            if not continue_copy:
                logging.info("Process terminated by user before copying files")
                print("Process terminated by user before copying files")
                return

            # Proceed to copy files
            logging.info("Copying files to destination folder")
            print("Copying files to destination folder")
            copy_files(df)
            logging.info("Process completed")
            print("Process completed")

            # Check if user wants to delete the copied files
            delete_files = CONFIG['behavior_flags']['auto_delete_copied_files']
            if not delete_files:
                delete_files_input = input("Do you want to delete the source files that were copied? (Y/N): ").strip().lower()
                delete_files = delete_files_input == 'y'
            if delete_files:
                logging.info("Deleting source files that were copied")
                print("Deleting source files that were copied")
                delete_copied_files(df)

            # Check if user wants to delete discarded files
            delete_discarded = CONFIG['behavior_flags']['auto_delete_discarded_files']
            if not delete_discarded:
                delete_discarded_input = input("Do you want to delete the discarded files? (Y/N): ").strip().lower()
                delete_discarded = delete_discarded_input == 'y'
            if delete_discarded:
                logging.info("Deleting discarded files")
                print("Deleting discarded files")
                delete_discarded_files(df)

            # Save updated metadata
            df.to_excel(metadata_file, index=False)
            logging.info(f"Updated metadata saved to {metadata_file}")
            print(f"Updated metadata saved to {metadata_file}")

            return
        else:
            print("No file selected. Exiting.")
            return

    # Check if hardcoded folders are valid
    if not (os.path.exists(source_folder) and os.path.isdir(source_folder)):
        source_folder = input("Enter the source folder path: ")
    if not (os.path.exists(dest_folder) and os.path.isdir(dest_folder)):
        dest_folder = input("Enter the destination folder path: ")

    # Detect if the destination folder is inside the source folder
    skip_dest_folder = False
    if os.path.commonpath([source_folder, dest_folder]) == source_folder:
        user_choice = input(f"The destination folder '{dest_folder}' is inside the source folder '{source_folder}'. Do you want to skip scanning the destination folder? (Y/N): ").strip().lower()
        if user_choice == 'y':
            skip_dest_folder = True

    logging.info("Starting the process")
    print("Starting the process")

    # Scan source folder and extract file metadata
    logging.info("Scanning source folder")
    print("Scanning source folder")
    df = process_files(source_folder, dest_folder, skip_dest_folder)

    # Log and print progress every 1000 files collected
    logging.info(f"{len(df)} files collected")
    print(f"{len(df)} files collected")

    if df.empty:
        logging.info("No files collected")
        print("No files collected")
        return

    # Prepare metadata for duplicate detection
    df = calculate_total_files_unified_day(df)
    df = calculate_time_intervals(df)
    df = calculate_short_sequence_counts(df)

    # Identify and mark duplicate files
    logging.info("Identifying duplicates")
    print("Identifying duplicates")
    df = mark_duplicates(df)

    # Display top unified days
    top_days = df['unified_day'].value_counts().head(20)
    for idx, (day, count) in enumerate(top_days.items(), start=1):
        log_msg = f"#{idx:02d} - {day} - {count} files"
        print(log_msg)
        logging.info(log_msg)

    # Ask for exclusion threshold
    default_exclude_threshold = CONFIG['processing_thresholds']['exclude_days_over_files']
    exclude_threshold_input = input(f"Do you want to exclude days with more than X files? Enter X value, or zero if you don't want to exclude any file (default: {default_exclude_threshold}): ").strip()
    exclude_threshold = int(exclude_threshold_input) if exclude_threshold_input else default_exclude_threshold
    if exclude_threshold > 0:
        df.loc[df['total_files_unified_day'] > exclude_threshold, 'exclude'] = 1

    # Ask for short sequence exclusion threshold
    default_short_seq_threshold = CONFIG['processing_thresholds']['exclude_short_sequences_over']
    short_seq_threshold_input = input(f"Enter the maximum number of files in short sequence to exclude days (default: {default_short_seq_threshold}): ").strip()
    short_seq_threshold = int(short_seq_threshold_input) if short_seq_threshold_input else default_short_seq_threshold
    if short_seq_threshold > 0:
        df.loc[df['short_sequence_count'] > short_seq_threshold, 'exclude'] = 1

    # Check to exclude files based on criteria
    exclude_criteria = CONFIG['behavior_flags']['exclude_creation_modified_criteria']
    if not exclude_criteria:
        exclude_criteria_input = input("Do you want to exclude files where the unified name is based on creation or modified dates? (Y/N): ").strip().lower()
        exclude_criteria = exclude_criteria_input == 'y'
    if exclude_criteria:
        df.loc[df['criteria'].isin(['creation', 'modified']), 'exclude'] = 1

    # Save inventory to Excel before copying
    logging.info("Saving inventory to Excel")
    print("Saving inventory to Excel")
    excel_path = os.path.join(dest_folder, f'metadata_{current_time_str}.xlsx')
    df.to_excel(excel_path, index=False)
    logging.info(f"Inventory saved to {excel_path}")
    print(f"Inventory saved to {excel_path}")

    # Check to continue with copying
    continue_copy = CONFIG['behavior_flags']['continue_copy']
    if not continue_copy:
        continue_copy_input = input("Do you want to continue with copying the files? (Y/N): ").strip().lower()
        continue_copy = continue_copy_input == 'y'
    if not continue_copy:
        logging.info("Process terminated by user before copying files")
        print("Process terminated by user before copying files")
        return

    # Copy qualifying files to destination with unified naming
    logging.info("Copying files to destination")
    print("Copying files to destination")
    copy_files(df)

    # Check if user wants to delete the copied files
    delete_files = CONFIG['behavior_flags']['auto_delete_copied_files']
    if not delete_files:
        delete_files_input = input("Do you want to delete the source files that were copied? (Y/N): ").strip().lower()
        delete_files = delete_files_input == 'y'
    if delete_files:
        logging.info("Deleting source files that were copied")
        print("Deleting source files that were copied")
        delete_copied_files(df)

    # Check if user wants to delete discarded files before copying
    delete_discarded = CONFIG['behavior_flags']['auto_delete_discarded_files']
    if not delete_discarded:
        delete_discarded_input = input("Do you want to delete the discarded files? (Y/N): ").strip().lower()
        delete_discarded = delete_discarded_input == 'y'
    if delete_discarded:
        logging.info("Deleting discarded files")
        print("Deleting discarded files")
        delete_discarded_files(df)

    # Save updated metadata
    df.to_excel(excel_path, index=False)
    logging.info(f"Updated metadata saved to {excel_path}")
    print(f"Updated metadata saved to {excel_path}")

    logging.info("Process completed")
    print("Process completed")

if __name__ == "__main__":
    try:
        source_folder = CONFIG['source_folder']
        dest_folder = CONFIG['destination_folder']
        main(source_folder, dest_folder)
    except KeyError as e:
        print(f"❌ Configuration error: Missing required configuration key: {e}")
        exit(1)
    except ValueError as e:
        print(f"❌ Configuration error: {e}")
        exit(1)
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        exit(1)