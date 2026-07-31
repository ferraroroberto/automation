"""
Duplicate detection and day-grouping for the photo-archive tool.

Everything that reasons over the metadata DataFrame rather than over files on
disk: unified-name collisions, SHA256 confirmation of true duplicates, per-day
file counts, and the burst ("short sequence") heuristic used to exclude
machine-gun shots. `recalculate_metadata` re-derives the whole lot from an
Excel inventory loaded back off disk.

Split out of `photos_archive.py` (audit issue #92).
"""

import hashlib
import logging
import os
from datetime import datetime
from typing import Optional

import pandas as pd

from photos_archive_config import get_config

# How often the SHA256 pass reports progress, in files.
_SHA256_PROGRESS_INTERVAL = 50


def calculate_time_intervals(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate time intervals between consecutive files within same day."""
    # Extract date part from unified names (YYYYMMDD)
    df['unified_day'] = df['unified_name'].apply(lambda x: x[:8])

    # Convert to datetime objects
    df['datetime'] = df['unified_name'].apply(lambda x: datetime.strptime(x[:15], '%Y%m%d-%H%M%S'))

    # Calculate seconds between consecutive files per day
    df['time_interval'] = df.groupby('unified_day')['datetime'].diff().dt.total_seconds().fillna(0)

    return df


def calculate_short_sequence_counts(df: pd.DataFrame, threshold: Optional[int] = None) -> pd.DataFrame:
    """Count files in short sequences per day (below time threshold)."""
    if threshold is None:
        threshold = get_config()['processing_thresholds']['short_sequence_seconds']

    # Ensure time intervals are calculated
    if 'time_interval' not in df.columns:
        df = calculate_time_intervals(df)

    # Mark files with short intervals
    df['short_sequence'] = df['time_interval'] <= threshold
    df['short_sequence_count'] = df.groupby('unified_day')['short_sequence'].transform('sum')

    # Clean up temporary columns
    df.drop(columns=['datetime', 'time_interval', 'short_sequence'], inplace=True)

    return df


def calculate_sha256(file_path: str) -> str:
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


def mark_duplicates(df: pd.DataFrame, check_sha256: Optional[bool] = None) -> pd.DataFrame:
    """
    Identifies and marks duplicate files in the DataFrame and counts duplicates.

    Args:
    df (pd.DataFrame): DataFrame containing file metadata.
    check_sha256 (Optional[bool]): Whether to confirm duplicates by hashing.
        None (the default) resolves it from config, prompting the user only if
        the config flag is off — pass an explicit bool to run unattended.

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
        if check_sha256 is None:
            check_sha256 = get_config()['duplicate_detection']['enable_sha256_check']
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
                if sha256_checked_count % _SHA256_PROGRESS_INTERVAL == 0:
                    logging.info("%d files checked for SHA256", sha256_checked_count)

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


def calculate_total_files_unified_day(df: pd.DataFrame) -> pd.DataFrame:
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


def log_top_unified_days(df: pd.DataFrame, limit: int = 20) -> None:
    """Log the days holding the most files, busiest first."""
    top_days = df['unified_day'].value_counts().head(limit)
    for idx, (day, count) in enumerate(top_days.items(), start=1):
        logging.info("#{:02d} - {} - {} files".format(idx, day, count))


def recalculate_metadata(df: pd.DataFrame) -> pd.DataFrame:
    """
    Recalculates the metadata fields based on the existing file information.

    Args:
    df (pd.DataFrame): DataFrame containing file metadata.

    Returns:
    pd.DataFrame: DataFrame with recalculated metadata fields.
    """
    destination_folder = get_config()['destination_folder']

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
            destination_folder) else 0,
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
    df['destination_path'] = df['unified_name'].apply(lambda x: os.path.join(destination_folder, x))

    # Display top unified days only once
    log_top_unified_days(df)

    return df
