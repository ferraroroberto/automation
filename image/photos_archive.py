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

This file is the CLI: the interactive prompts, the copy/delete file operations
they gate, and the two run flows (existing metadata vs fresh scan). The logic
underneath lives in three siblings, split out by audit issue #92 —
`photos_archive_config.py` (config + logging), `photos_archive_metadata.py`
(date extraction + scanning), `photos_archive_duplicates.py` (duplicate and
day-grouping analysis).
"""

import logging
import os
import shutil
import tkinter as tk
from datetime import datetime
from tkinter import filedialog
from typing import Optional

import pandas as pd

from photos_archive_config import get_config, setup_logging
from photos_archive_duplicates import (
    calculate_short_sequence_counts,
    calculate_time_intervals,
    calculate_total_files_unified_day,
    log_top_unified_days,
    mark_duplicates,
    recalculate_metadata,
)
from photos_archive_metadata import process_files

# How often the copy/delete passes report progress, in files.
_FILE_OP_PROGRESS_INTERVAL = 1000


def _prompt_yes_no(config_flag: bool, question: str) -> bool:
    """
    Resolve a yes/no decision: honour the config flag when it's set, otherwise ask.

    Every `behavior_flags` entry works this way — True means "don't ask, just do
    it", False means "ask me at runtime".
    """
    if config_flag:
        return True
    return input(question).strip().lower() == 'y'


def _delete_flagged_files(df: pd.DataFrame, mask: pd.Series, log_label: str) -> None:
    """
    Delete the source files selected by `mask`, marking df['deleted'] in place.

    Shared by delete_discarded_files and delete_copied_files, which differ
    only in which rows they select for deletion.

    Args:
        df (pd.DataFrame): DataFrame containing file metadata (needs
            'file_path' and 'deleted' columns).
        mask (pd.Series): Boolean Series (aligned to df's index) selecting
            which rows to attempt deletion for.
        log_label (str): Noun phrase used in progress/summary log messages,
            e.g. "discarded files" or "files".
    """
    deleted_files_count = 0
    last_logged_count = 0

    for index, row in df.iterrows():
        if mask.loc[index]:
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

        if deleted_files_count // _FILE_OP_PROGRESS_INTERVAL > last_logged_count // _FILE_OP_PROGRESS_INTERVAL:
            logging.info("%d %s deleted", deleted_files_count, log_label)
            last_logged_count = deleted_files_count

    logging.info("%d %s deleted", deleted_files_count, log_label)


def delete_discarded_files(df: pd.DataFrame) -> None:
    """
    Deletes the source files that are marked as discarded in the metadata.

    Args:
    df (pd.DataFrame): DataFrame containing file metadata.
    """
    _delete_flagged_files(df, (df['deleted'] == 0) & (df['discard'] == 1), "discarded files")


def delete_copied_files(df: pd.DataFrame) -> None:
    """
    Deletes the source files that were successfully copied to the destination folder.

    Args:
    df (pd.DataFrame): DataFrame containing file metadata.
    """
    _delete_flagged_files(df, df['copy_success'] == 1, "files")


def copy_files(df: pd.DataFrame) -> None:
    """
    Copies non-duplicate files to the destination folder, renaming them based on the unified name.

    Args:
    df (pd.DataFrame): DataFrame containing file metadata.
    """
    destination_folder = get_config()['destination_folder']
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

            dest_path = os.path.join(destination_folder, new_file_name)
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
        if copied_files_count // _FILE_OP_PROGRESS_INTERVAL > last_logged_count // _FILE_OP_PROGRESS_INTERVAL:
            logging.info("%d files copied", copied_files_count)
            last_logged_count = copied_files_count

    logging.info("%d files copied", copied_files_count)


def _apply_exclusion_thresholds(df: pd.DataFrame) -> None:
    """
    Display the top unified days and prompt for the exclude-days-over-X,
    short-sequence, and creation/modified-criteria exclusion thresholds,
    mutating df['exclude'] in place.

    Shared by main()'s "use existing metadata + recalculate" and "fresh scan"
    flows (dedup: audit issue #64).
    """
    config = get_config()
    log_top_unified_days(df)

    # Ask for exclusion threshold
    default_exclude_threshold = config['processing_thresholds']['exclude_days_over_files']
    exclude_threshold_input = input(f"Do you want to exclude days with more than X files? Enter X value, or zero if you don't want to exclude any file (default: {default_exclude_threshold}): ").strip()
    exclude_threshold = int(exclude_threshold_input) if exclude_threshold_input else default_exclude_threshold
    if exclude_threshold > 0:
        df.loc[df['total_files_unified_day'] > exclude_threshold, 'exclude'] = 1

    # Ask for short sequence exclusion threshold
    default_short_seq_threshold = config['processing_thresholds']['exclude_short_sequences_over']
    short_seq_threshold_input = input(f"Enter the maximum number of files in short sequence to exclude days (default: {default_short_seq_threshold}): ").strip()
    short_seq_threshold = int(short_seq_threshold_input) if short_seq_threshold_input else default_short_seq_threshold
    if short_seq_threshold > 0:
        df.loc[df['short_sequence_count'] > short_seq_threshold, 'exclude'] = 1

    # Check to exclude files based on criteria
    exclude_criteria = _prompt_yes_no(
        config['behavior_flags']['exclude_creation_modified_criteria'],
        "Do you want to exclude files where the unified name is based on creation or modified dates? (Y/N): ",
    )
    if exclude_criteria:
        df.loc[df['criteria'].isin(['creation', 'modified']), 'exclude'] = 1


def _confirm_copy_and_cleanup(df: pd.DataFrame, metadata_path: str) -> None:
    """
    Prompt to continue with copying, copy qualifying files, optionally delete
    copied/discarded source files, then save the (possibly mutated) metadata
    back to metadata_path.

    Shared by main()'s "use existing metadata" and "fresh scan" flows (dedup:
    audit issue #64).
    """
    behavior_flags = get_config()['behavior_flags']

    # Check for confirmation to process files
    if not _prompt_yes_no(behavior_flags['continue_copy'],
                          "Do you want to continue with copying the files? (Y/N): "):
        logging.info("Process terminated by user before copying files")
        return

    # Proceed to copy files
    logging.info("Copying files to destination folder")
    copy_files(df)

    # Check if user wants to delete the copied files
    if _prompt_yes_no(behavior_flags['auto_delete_copied_files'],
                      "Do you want to delete the source files that were copied? (Y/N): "):
        logging.info("Deleting source files that were copied")
        delete_copied_files(df)

    # Check if user wants to delete discarded files
    if _prompt_yes_no(behavior_flags['auto_delete_discarded_files'],
                      "Do you want to delete the discarded files? (Y/N): "):
        logging.info("Deleting discarded files")
        delete_discarded_files(df)

    # Save updated metadata
    df.to_excel(metadata_path, index=False)
    logging.info("Updated metadata saved to %s", metadata_path)
    logging.info("Process completed")


def _ask_for_metadata_file() -> Optional[str]:
    """Open a centered Tkinter file dialog for an existing metadata workbook."""
    # Ensure Tkinter is properly initialized
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Get the screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Set the geometry of the root window to the center of the screen
    window_width = 300  # Width of the dialog window
    window_height = 200  # Height of the dialog window
    position_right = int(screen_width / 2 - window_width / 2)
    position_down = int(screen_height / 2 - window_height / 2)

    # Adjust the geometry
    root.geometry(f"{window_width}x{window_height}+{position_right}+{position_down}")

    # Open the file dialog
    metadata_file = filedialog.askopenfilename(title="Select metadata Excel file", filetypes=[("Excel files", "*.xlsx")])
    root.destroy()  # Properly destroy the Tkinter root window
    return metadata_file or None


def _run_from_existing_metadata(dest_folder: str) -> None:
    """
    Prompt for an existing metadata Excel file via a Tkinter dialog, optionally
    recalculate the duplicate/exclusion logic, then confirm/copy/cleanup.

    Split out of main()'s "use existing metadata" branch (audit issue #67).
    """
    metadata_file = _ask_for_metadata_file()
    if not metadata_file:
        logging.info("No file selected. Exiting.")
        return

    df = pd.read_excel(metadata_file)
    logging.info("Loaded metadata from %s", metadata_file)

    # Check if user wants to recalculate the logic
    if _prompt_yes_no(get_config()['behavior_flags']['recalculate_logic'],
                      "Do you want to recalculate the logic? (Y/N): "):
        df = recalculate_metadata(df)
        _apply_exclusion_thresholds(df)

        current_time_str = datetime.now().strftime('%Y%m%d-%H%M')
        metadata_file = os.path.join(dest_folder, f'metadata_{current_time_str}.xlsx')
        df.to_excel(metadata_file, index=False)
        logging.info("Recalculated metadata and saved to %s", metadata_file)

    _confirm_copy_and_cleanup(df, metadata_file)


def _run_fresh_scan(source_folder: str, dest_folder: str, current_time_str: str) -> None:
    """
    Scan source_folder from scratch, compute duplicate/exclusion metadata,
    save an inventory Excel, then confirm/copy/cleanup.

    Split out of main()'s "fresh scan" branch (audit issue #67).
    """
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

    # Scan source folder and extract file metadata
    logging.info("Scanning source folder")
    df = process_files(source_folder, dest_folder, skip_dest_folder)

    logging.info("%d files collected", len(df))

    if df.empty:
        logging.info("No files collected")
        return

    # Prepare metadata for duplicate detection
    df = calculate_total_files_unified_day(df)
    df = calculate_time_intervals(df)
    df = calculate_short_sequence_counts(df)

    # Identify and mark duplicate files
    logging.info("Identifying duplicates")
    df = mark_duplicates(df)

    _apply_exclusion_thresholds(df)

    # Save inventory to Excel before copying
    logging.info("Saving inventory to Excel")
    excel_path = os.path.join(dest_folder, f'metadata_{current_time_str}.xlsx')
    df.to_excel(excel_path, index=False)
    logging.info("Inventory saved to %s", excel_path)

    _confirm_copy_and_cleanup(df, excel_path)


def main(source_folder: str, dest_folder: str) -> None:
    """Run the archive: either replay an existing inventory or scan from scratch."""
    # Set up logging to save log file in the destination folder
    current_time_str = datetime.now().strftime('%Y%m%d-%H%M')
    setup_logging(dest_folder)

    # Check if user wants to use existing metadata
    if _prompt_yes_no(get_config()['behavior_flags']['use_existing_metadata'],
                      "Do you want to use an existing metadata file? (Y/N): "):
        _run_from_existing_metadata(dest_folder)
    else:
        _run_fresh_scan(source_folder, dest_folder, current_time_str)


if __name__ == "__main__":
    try:
        config = get_config()
        main(config['source_folder'], config['destination_folder'])
    except KeyError as e:
        logging.error("❌ Configuration error: Missing required configuration key: %s", e)
        exit(1)
    except ValueError as e:
        logging.error("❌ Configuration error: %s", e)
        exit(1)
    except Exception as e:
        logging.error("❌ Unexpected error: %s", e)
        exit(1)
