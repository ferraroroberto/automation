"""
Media metadata extraction for the photo-archive tool.

Answers one question per file: when was it created? Tries EXIF
`DateTimeOriginal`, then video container metadata, then the configured
filename patterns, and finally falls back to the filesystem's creation /
modified timestamps. `get_file_info` turns that into the one row the rest of
the pipeline works on; `process_files` walks a tree into a DataFrame.

Split out of `photos_archive.py` (audit issue #92).
"""

import logging
import os
import re
from datetime import datetime, timedelta
from typing import Any, Dict, Optional

import pandas as pd
from PIL import Image
from PIL.ExifTags import TAGS
from pymediainfo import MediaInfo

from photos_archive_config import get_config

# How often the directory walk reports progress, in files.
_SCAN_PROGRESS_INTERVAL = 1000


def get_exif_creation_date(file_path: str) -> Optional[datetime]:
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


def parse_date(date_str: str) -> Optional[datetime]:
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


def extract_date_from_filename(filename: str) -> Optional[datetime]:
    """Extract date from filename using configured regex patterns."""
    patterns = get_config()['date_parsing']['filename_patterns']
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


def get_video_creation_date(file_path: str) -> Optional[datetime]:
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


def get_file_info(file_path: str) -> Optional[Dict[str, Any]]:
    """Extract comprehensive metadata from media file."""
    config = get_config()
    try:
        file_stat = os.stat(file_path)
        creation_time = datetime.fromtimestamp(file_stat.st_ctime)
        modified_time = datetime.fromtimestamp(file_stat.st_mtime)

        # Initialize defaults
        file_type = 'Unknown'
        exif_date = None
        video_date = None
        criteria = "modified"

        # Determine file type and extract creation date
        file_extension_lower = file_path.lower()
        image_extensions = [ext.lower() for ext in config['file_extensions']['image']]
        video_extensions = [ext.lower() for ext in config['file_extensions']['video']]

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
        dest_path = os.path.join(config['destination_folder'], unified_name)

        # Check if the file is in the destination directory (but not in its subfolders)
        if os.path.abspath(os.path.dirname(file_path)) == os.path.abspath(config['destination_folder']):
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
            'exclude': exclude
        }
    except Exception as e:
        # Log error if file information cannot be retrieved
        logging.error(f"Error getting file info for {file_path}: {e}")
        return None


def process_files(source_folder: str, dest_folder: str, skip_dest_folder: bool) -> pd.DataFrame:
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
            if len(file_records) % _SCAN_PROGRESS_INTERVAL == 0:
                logging.info("%d files processed", len(file_records))

    return pd.DataFrame(file_records)
