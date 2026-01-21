#!/usr/bin/env python3
"""
CSV Structure Comparison and Processing Tool

Compares the structure (column names, order, and data types) of two CSV files
and generates a detailed report of any changes. Additionally processes the
new CSV file by filtering columns and dates, then saves the filtered data
to Excel format.

The script will prompt for confirmation before processing the CSV file.

Usage:
    python csv_structure_compare.py [directory_path] [--start-date YYYY-MM-DD] [--end-date YYYY-MM-DD]

If no directory is provided, uses the default from csv_compare_config.json
Date filtering defaults to the last closed quarter if not specified
"""

import json
import logging
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Load configuration
CONFIG_FILE = Path(__file__).parent / "csv_compare_config.json"

def load_config() -> Dict:
    """Load configuration from JSON file."""
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
        logger.info("📂 Configuration loaded successfully")
        return config
    except FileNotFoundError:
        logger.error(f"❌ Configuration file not found: {CONFIG_FILE}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        logger.error(f"❌ Invalid JSON in config file: {e}")
        sys.exit(1)

def read_csv_structure(file_path: Path) -> Tuple[List[str], Dict[str, str]]:
    """
    Read CSV file and extract structure information.

    Returns:
        Tuple of (column_names, column_types_dict)
    """
    try:
        # Read first few rows to determine structure
        df = pd.read_csv(file_path, nrows=5)

        column_names = df.columns.tolist()
        column_types = {col: str(df[col].dtype) for col in column_names}

        logger.info(f"📊 Read {len(column_names)} columns from {file_path.name}")
        return column_names, column_types

    except FileNotFoundError:
        logger.error(f"❌ File not found: {file_path}")
        return [], {}
    except Exception as e:
        logger.error(f"❌ Error reading CSV {file_path}: {e}")
        return [], {}

def compare_csv_structures(
    old_columns: List[str],
    old_types: Dict[str, str],
    new_columns: List[str],
    new_types: Dict[str, str]
) -> Dict[str, List[str]]:
    """
    Compare two CSV structures and identify differences.

    Returns:
        Dictionary with change categories and lists of affected columns
    """
    changes = {
        "added_columns": [],
        "removed_columns": [],
        "order_changed": [],
        "type_changed": []
    }

    # Find added and removed columns
    old_set = set(old_columns)
    new_set = set(new_columns)

    changes["added_columns"] = sorted(list(new_set - old_set))
    changes["removed_columns"] = sorted(list(old_set - new_set))

    # Check column order for common columns
    common_columns = old_set & new_set
    old_order = {col: idx for idx, col in enumerate(old_columns)}
    new_order = {col: idx for idx, col in enumerate(new_columns)}

    for col in common_columns:
        if old_order[col] != new_order[col]:
            changes["order_changed"].append(col)

    # Check type changes for common columns
    for col in common_columns:
        if old_types.get(col) != new_types.get(col):
            changes["type_changed"].append(
                f"{col}: {old_types.get(col, 'unknown')} -> {new_types.get(col, 'unknown')}"
            )

    return changes

def generate_report(
    old_file: str,
    new_file: str,
    changes: Dict[str, List[str]]
) -> str:
    """Generate a formatted report of the comparison results."""
    report_lines = [
        "=" * 60,
        "CSV STRUCTURE COMPARISON REPORT",
        "=" * 60,
        f"Old file: {old_file}",
        f"New file: {new_file}",
        ""
    ]

    has_changes = any(changes.values())

    if not has_changes:
        report_lines.append("[OK] No structural changes detected")
        report_lines.append("")
        report_lines.append("Both files have identical column structure.")
    else:
        report_lines.append("[WARNING] Structural changes detected:")
        report_lines.append("")

        for change_type, items in changes.items():
            if items:
                # Format change type for display
                display_name = change_type.replace("_", " ").title()
                report_lines.append(f"• {display_name}:")
                for item in items:
                    report_lines.append(f"   - {item}")
                report_lines.append("")

    report_lines.append("=" * 60)
    return "\n".join(report_lines)

def filter_columns(df: pd.DataFrame, columns_to_keep: List[str]) -> pd.DataFrame:
    """
    Filter DataFrame to keep only specified columns.

    Args:
        df: Input DataFrame
        columns_to_keep: List of column names to keep

    Returns:
        Filtered DataFrame
    """
    available_columns = [col for col in columns_to_keep if col in df.columns]
    missing_columns = [col for col in columns_to_keep if col not in df.columns]

    if missing_columns:
        logger.warning(f"⚠️  Missing columns: {missing_columns}")

    if not available_columns:
        logger.error("❌ None of the specified columns found in DataFrame")
        return df

    logger.info(f"📊 Keeping {len(available_columns)} columns: {available_columns}")
    return df[available_columns]

def get_last_closed_quarter_dates(current_date: Optional[datetime] = None) -> Tuple[datetime, datetime]:
    """
    Calculate the start and end dates of the last closed quarter.

    Args:
        current_date: Reference date (defaults to today)

    Returns:
        Tuple of (start_date, end_date) for the last closed quarter
    """
    if current_date is None:
        current_date = datetime.now()

    # Get the current quarter
    current_quarter = ((current_date.month - 1) // 3) + 1
    current_year = current_date.year

    # Calculate last closed quarter
    if current_quarter == 1:
        # If we're in Q1, last closed quarter is Q4 of previous year
        quarter = 4
        year = current_year - 1
    else:
        # Otherwise, last quarter of current year
        quarter = current_quarter - 1
        year = current_year

    # Calculate start and end dates
    start_month = (quarter - 1) * 3 + 1
    end_month = quarter * 3

    start_date = datetime(year, start_month, 1)
    # End date is the last day of the quarter
    if end_month == 12:
        end_date = datetime(year, 12, 31)
    else:
        end_date = datetime(year, end_month + 1, 1) - timedelta(days=1)

    logger.info(f"📅 Last closed quarter: Q{quarter} {year} ({start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')})")
    return start_date, end_date

def filter_by_date_range(
    df: pd.DataFrame,
    date_column: str,
    start_date: Optional[datetime] = None,
    end_date: Optional[datetime] = None,
    use_last_quarter: bool = True
) -> pd.DataFrame:
    """
    Filter DataFrame by date range.

    Args:
        df: Input DataFrame
        date_column: Name of the date column
        start_date: Start date for filtering (optional)
        end_date: End date for filtering (optional)
        use_last_quarter: Whether to use last closed quarter as default

    Returns:
        Filtered DataFrame
    """
    if date_column not in df.columns:
        logger.warning(f"⚠️  Date column '{date_column}' not found. Skipping date filtering.")
        return df

    # Convert date column to datetime if it's not already
    try:
        df[date_column] = pd.to_datetime(df[date_column], errors='coerce')
    except Exception as e:
        logger.error(f"❌ Error converting date column: {e}")
        return df

    # Determine date range
    if start_date is None and end_date is None and use_last_quarter:
        start_date, end_date = get_last_closed_quarter_dates()

    if start_date is None or end_date is None:
        logger.info("📅 No date filtering applied")
        return df

    # Apply date filter
    original_count = len(df)
    mask = (df[date_column] >= start_date) & (df[date_column] <= end_date)
    filtered_df = df[mask].copy()

    filtered_count = len(filtered_df)
    logger.info(f"📅 Filtered {original_count} rows to {filtered_count} rows "
                f"(date range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')})")

    return filtered_df

def save_to_excel(df: pd.DataFrame, output_path: Path, sheet_name: str = "Filtered Data") -> bool:
    """
    Save DataFrame to Excel file.

    Args:
        df: DataFrame to save
        output_path: Path to save the Excel file
        sheet_name: Name of the Excel sheet

    Returns:
        True if successful, False otherwise
    """
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        logger.info(f"💾 Saved {len(df)} rows to {output_path}")
        return True

    except Exception as e:
        logger.error(f"❌ Error saving to Excel: {e}")
        return False

def process_csv_file(
    file_path: Path,
    config: Dict,
    start_date: Optional[datetime] = None,
    end_date: Optional[datetime] = None
) -> Optional[pd.DataFrame]:
    """
    Process a single CSV file: load, filter columns, filter by date.

    Args:
        file_path: Path to the CSV file
        config: Configuration dictionary
        start_date: Custom start date (optional)
        end_date: Custom end date (optional)

    Returns:
        Processed DataFrame or None if error
    """
    try:
        logger.info(f"📖 Processing {file_path.name}...")

        # Read CSV
        df = pd.read_csv(file_path)
        logger.info(f"📊 Loaded {len(df)} rows, {len(df.columns)} columns")

        # Filter columns
        if "columns_to_keep" in config and config["columns_to_keep"]:
            df = filter_columns(df, config["columns_to_keep"])

        # Filter by date
        if "date_column" in config and "filter_by_quarter" in config:
            df = filter_by_date_range(
                df,
                config["date_column"],
                start_date,
                end_date,
                config["filter_by_quarter"]
            )

        logger.info(f"✅ Processed {len(df)} rows, {len(df.columns)} columns")
        return df

    except Exception as e:
        logger.error(f"❌ Error processing {file_path}: {e}")
        return None

def main(
    directory_path: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None
):
    """Main function to run the CSV comparison and processing."""
    # Load configuration
    config = load_config()

    # Determine directory to use
    if directory_path:
        target_dir = Path(directory_path)
    else:
        target_dir = Path(config["default_directory"])

    logger.info(f"🎯 Using directory: {target_dir}")

    # Check if directory exists
    if not target_dir.exists():
        logger.error(f"❌ Directory not found: {target_dir}")
        sys.exit(1)

    # Parse date parameters
    start_dt = None
    end_dt = None
    if start_date:
        try:
            start_dt = datetime.fromisoformat(start_date)
        except ValueError:
            logger.error(f"❌ Invalid start date format: {start_date}. Use YYYY-MM-DD")
            sys.exit(1)
    if end_date:
        try:
            end_dt = datetime.fromisoformat(end_date)
        except ValueError:
            logger.error(f"❌ Invalid end date format: {end_date}. Use YYYY-MM-DD")
            sys.exit(1)

    # Define file paths
    old_file = target_dir / config["file_old"]
    new_file = target_dir / config["file_new"]

    # Read both CSV structures
    logger.info("📖 Reading old CSV structure...")
    old_columns, old_types = read_csv_structure(old_file)

    logger.info("📖 Reading new CSV structure...")
    new_columns, new_types = read_csv_structure(new_file)

    # Check if both files were read successfully
    if not old_columns and not new_columns:
        logger.error("❌ Both CSV files could not be read")
        sys.exit(1)
    elif not old_columns:
        logger.error(f"❌ Old CSV file could not be read: {old_file}")
        sys.exit(1)
    elif not new_columns:
        logger.error(f"❌ New CSV file could not be read: {new_file}")
        sys.exit(1)

    # Compare structures
    logger.info("🔍 Comparing structures...")
    changes = compare_csv_structures(old_columns, old_types, new_columns, new_types)

    # Generate and print report
    report = generate_report(config["file_old"], config["file_new"], changes)
    print("\n" + report)

    # Log summary
    has_changes = any(changes.values())
    if has_changes:
        logger.warning("⚠️  Structural changes detected")
    else:
        logger.info("✅ No structural changes detected")

    # Confirm before processing
    print("\n" + "="*60)
    try:
        confirm = input("Ready to process and filter the new CSV file? (y/N): ").strip().lower()
        if confirm not in ['y', 'yes']:
            logger.info("⏹️  Processing cancelled by user")
            print("Processing cancelled.")
            return
    except (KeyboardInterrupt, EOFError):
        logger.info("⏹️  Processing cancelled by user")
        print("\nProcessing cancelled.")
        return

    print()  # Line break
    logger.info("🔄 Processing and filtering new CSV file...")
    processed_df = process_csv_file(new_file, config, start_dt, end_dt)

    if processed_df is not None and not processed_df.empty:
        # Save to Excel
        output_file = target_dir / config.get("output_file", "filtered_payments.xlsx")
        logger.info(f"💾 Saving filtered data to {output_file}...")

        if save_to_excel(processed_df, output_file):
            logger.info("✅ Processing completed successfully")
            print(f"\n[OK] Filtered data saved to: {output_file}")
            print(f"   Rows: {len(processed_df)}, Columns: {len(processed_df.columns)}")
        else:
            logger.error("❌ Failed to save filtered data")
            sys.exit(1)
    else:
        logger.error("❌ No data to save after processing")
        sys.exit(1)

if __name__ == "__main__":
    # Parse command line arguments
    import argparse

    parser = argparse.ArgumentParser(description="CSV Structure Comparison and Processing Tool")
    parser.add_argument("directory_path", nargs="?", help="Directory containing CSV files (optional)")
    parser.add_argument("--start-date", help="Start date for filtering (YYYY-MM-DD)")
    parser.add_argument("--end-date", help="End date for filtering (YYYY-MM-DD)")

    args = parser.parse_args()

    main(args.directory_path, args.start_date, args.end_date)