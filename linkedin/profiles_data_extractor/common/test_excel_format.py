"""
Test script for Excel Format Manager - Parts 1, 2 & 3

Tests:
1. Get format (save Excel format to JSON)
2. Copy file and apply format
3. Convert URL columns to hyperlinks
"""

import json
import logging
import os
from pathlib import Path

import pandas as pd
from excel_format_manager import (
    save_excel_format_to_json,
    convert_url_columns_to_hyperlinks,
    apply_format_from_json
)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def wait_for_enter() -> None:
    """Wait for user to press Enter key."""
    logger.info("")
    logger.info("=" * 80)
    logger.info("⏸️  PAUSED: Check the raw format in Excel")
    logger.info("=" * 80)
    logger.info("Press ENTER to continue and apply formatting...")
    input()


def main() -> None:
    """Test parts 1, 2 and 3 of Excel format manager."""
    # Load configuration
    config_path = os.path.join(
        os.path.dirname(__file__),
        "linkedin_profiles_data.json"
    )
    
    if not os.path.exists(config_path):
        logger.error(f"❌ Config file not found: {config_path}")
        return
    
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    destination_file = config.get("destination_file")
    if not destination_file:
        logger.error("❌ 'destination_file' not found in config")
        return
    
    # Create format JSON path (same directory as config)
    format_json_path = os.path.join(
        os.path.dirname(config_path),
        "excel_format_spec.json"
    )
    
    logger.info("=" * 80)
    logger.info("🧪 Testing Excel Format Manager - Parts 1, 2 & 3")
    logger.info("=" * 80)
    logger.info(f"📄 Excel file: {destination_file}")
    logger.info(f"📋 Format JSON: {format_json_path}")
    logger.info("")
    
    # Part 1: Save format to JSON
    logger.info("=" * 80)
    logger.info("PART 1: Saving Excel format to JSON")
    logger.info("=" * 80)
    if os.path.exists(destination_file):
        success = save_excel_format_to_json(destination_file, format_json_path)
        if success:
            logger.info("✅ Part 1 completed successfully")
        else:
            logger.error("❌ Part 1 failed")
    else:
        logger.error(f"❌ Excel file does not exist: {destination_file}")
        return
    
    logger.info("")
    
    # Part 2: Copy file and apply format
    logger.info("=" * 80)
    logger.info("PART 2: Copy file and apply format")
    logger.info("=" * 80)
    
    # Create test file path with _test suffix
    dest_path = Path(destination_file)
    test_file_path = dest_path.parent / f"{dest_path.stem}_test{dest_path.suffix}"
    
    logger.info(f"📄 Test file will be: {test_file_path}")
    logger.info("")
    
    # Step 1: Read dataframe from destination_file
    logger.info("STEP 2.1: Reading dataframe from source file")
    try:
        df = pd.read_excel(destination_file, engine='openpyxl')
        logger.info(f"✅ Loaded dataframe with {len(df)} rows and {len(df.columns)} columns")
        logger.info(f"   Columns: {', '.join(df.columns.tolist())}")
    except Exception as e:
        logger.error(f"❌ Error reading Excel file: {e}")
        return
    
    logger.info("")
    
    # Step 2: Export to new Excel file with _test suffix
    logger.info("STEP 2.2: Exporting to test file (raw format, no formatting)")
    try:
        df.to_excel(test_file_path, index=False, engine='openpyxl')
        logger.info(f"✅ Exported to: {test_file_path}")
        logger.info("   File is now ready - check the raw format (no formatting applied)")
    except Exception as e:
        logger.error(f"❌ Error exporting Excel file: {e}")
        return
    
    logger.info("")
    
    # Step 3: Wait for Enter keypress
    wait_for_enter()
    
    # Step 4: Apply format from JSON
    logger.info("")
    logger.info("STEP 2.3: Applying format from JSON")
    success = apply_format_from_json(str(test_file_path), format_json_path)
    
    if success:
        logger.info("✅ Format applied successfully")
        logger.info(f"   Check the formatted file: {test_file_path}")
    else:
        logger.error("❌ Failed to apply format")
    
    logger.info("")
    
    # Part 3: Convert URL columns to hyperlinks
    logger.info("=" * 80)
    logger.info("PART 3: Converting URL columns to hyperlinks")
    logger.info("=" * 80)
    success = convert_url_columns_to_hyperlinks(str(test_file_path))
    if success:
        logger.info("✅ Part 3 completed successfully")
    else:
        logger.error("❌ Part 3 failed")
    
    logger.info("")
    logger.info("=" * 80)
    logger.info("✅ Testing complete")
    logger.info("=" * 80)
    logger.info(f"📄 Final test file: {test_file_path}")


if __name__ == "__main__":
    main()

