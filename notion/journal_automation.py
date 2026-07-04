"""Notion journal automation: weekly consolidated export for LLM analysis."""

import logging
import os
import sys
from datetime import datetime, timedelta
from typing import Any, Optional

from utils import load_env_variables, load_json_config
import utils as notion_utils

logger = logging.getLogger(__name__)

_ADJUSTED_DATE: Optional[datetime] = None
_NON_INTERACTIVE = False

def _cli_timedelta() -> Optional[int]:
    """Return the integer offset from argv[1] if present and valid, else None."""
    if len(sys.argv) < 2:
        return None
    try:
        return int(sys.argv[1])
    except ValueError:
        return None

def get_adjusted_date(timedelta_value: Optional[int] = None) -> datetime:
    """Return the adjusted date for processing (cached after first call)."""
    global _ADJUSTED_DATE, _NON_INTERACTIVE

    if _ADJUSTED_DATE is not None:
        return _ADJUSTED_DATE

    current_date = datetime.now()
    if timedelta_value is None:
        cli_value = _cli_timedelta()
        if cli_value is not None:
            timedelta_value = cli_value
            _NON_INTERACTIVE = True
        else:
            raw = input("📅 Enter an integer number to apply a timedelta to the current date (default is 0): ")
            try:
                timedelta_value = int(raw)
            except ValueError:
                timedelta_value = 0

    _ADJUSTED_DATE = current_date + timedelta(days=timedelta_value)
    return _ADJUSTED_DATE

def calculate_week_range(current_date: datetime) -> tuple[datetime, datetime]:
    """
    Return (start_date, end_date) for the previous week (Monday–Sunday).
    If today is Sunday, end_date is today.
    """
    if current_date.weekday() == 6:
        end_date = current_date
    else:
        days_since_sunday = (current_date.weekday() + 1) % 7 or 7
        end_date = current_date - timedelta(days=days_since_sunday)

    start_date = end_date - timedelta(days=6)
    start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)
    return start_date, end_date

def query_notion_database(
    database_id: str,
    headers: dict[str, str],
    start_date: datetime,
    end_date: datetime,
    date_property: str,
) -> list[dict[str, Any]]:
    """Query Notion database for journal entries in the date range."""
    logger.info("🔍 Querying Notion database...")

    start_iso = start_date.isoformat()
    end_iso = end_date.isoformat()

    query_body = {
        "filter": {
            "and": [
                {
                    "property": date_property,
                    "date": {
                        "on_or_after": start_iso
                    }
                },
                {
                    "property": date_property,
                    "date": {
                        "on_or_before": end_iso
                    }
                }
            ]
        },
        "sorts": [
            {
                "property": date_property,
                "direction": "ascending"
            }
        ]
    }

    return notion_utils.paginated_database_query(headers, database_id, query_body)

# Property types this extractor knows how to render as a display string.
# Anything else (e.g. url, number, relation) falls through to "" - matches
# the pre-dedup behavior, which only ever handled these six types.
_HANDLED_PROPERTY_TYPES = {'title', 'rich_text', 'multi_select', 'select', 'checkbox', 'date'}


def extract_property_value(properties: dict[str, Any], property_name: str) -> str:
    """Extract string value from a Notion property, or empty string if missing."""
    prop = properties.get(property_name, {})
    prop_type = prop.get('type')

    if prop_type not in _HANDLED_PROPERTY_TYPES:
        return ""

    value = notion_utils.extract_property(prop)

    if isinstance(value, bool):
        return "✓" if value else ""
    if isinstance(value, list):
        return ', '.join(value)
    return value or ""

def process_field_from_pages(pages: list[dict], field_config: dict[str, Any]) -> str:
    """Process one field from Notion pages; return formatted string with frequency counts."""
    column_name = field_config['column']
    split_comma = field_config.get('split_comma', False)
    
    # Extract all values for this field and count frequencies
    value_counts = {}
    for page in pages:
        properties = page.get('properties', {})
        value = extract_property_value(properties, column_name)
        
        if value and value != '[]':
            if split_comma:
                # Split on comma and add individual items
                items = [item.strip() for item in value.split(',')]
                for item in items:
                    if item:
                        value_counts[item] = value_counts.get(item, 0) + 1
            else:
                value_counts[value] = value_counts.get(value, 0) + 1
    
    # Format output with frequency counts
    output_lines = []
    for value, count in value_counts.items():
        if count > 1:
            output_lines.append(f"{value} ({count}x)")
        else:
            output_lines.append(value)
    
    # Join with newlines
    return '\n'.join(output_lines)

def process_checkboxes_from_pages(
    pages: list[dict], checkbox_configs: list[dict[str, Any]]
) -> str:
    """Count checked days per checkbox; return summary string (e.g. 'patience: 5 days out of 7')."""
    total_days = len(pages)
    checkbox_counts = {}
    
    # Count checked days for each checkbox
    for checkbox_config in checkbox_configs:
        column_name = checkbox_config['column']
        checkbox_counts[column_name] = 0
        
        for page in pages:
            properties = page.get('properties', {})
            value = extract_property_value(properties, column_name)
            
            # If checkbox is checked, value will be "✓"
            if value == "✓":
                checkbox_counts[column_name] += 1
    
    # Format output
    output_lines = []
    for checkbox_config in checkbox_configs:
        column_name = checkbox_config['column']
        display_name = checkbox_config.get('display_name', column_name)
        count = checkbox_counts.get(column_name, 0)
        output_lines.append(f"{display_name}: {count} days out of {total_days}")
    
    return '\n'.join(output_lines)

def _open_folder_in_explorer(folder_path: str) -> None:
    """Open the given folder in the system file manager (Windows Explorer)."""
    if os.name != "nt":
        return
    path = os.path.abspath(folder_path)
    if not os.path.isdir(path):
        return
    os.startfile(path)


def process_journal_consolidated(config: dict[str, Any], env_vars: dict[str, str]) -> Optional[str]:
    """Process Notion journal entries and write consolidated output file. Returns output path or None."""
    current_date = get_adjusted_date()
    start_date, end_date = calculate_week_range(current_date)

    logger.info(
        "📅 Date range: %s to %s",
        start_date.strftime("%Y-%m-%d"),
        end_date.strftime("%Y-%m-%d"),
    )
    
    notion_api_key = env_vars.get("notion_api_token")
    if not notion_api_key:
        raise ValueError("NOTION_API_TOKEN not found in environment variables")
    
    headers = {
        "Authorization": f"Bearer {notion_api_key}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json",
    }

    database_id = config["input"]["database_id"]
    date_property = config["input"].get("date_property", "Date")
    
    pages = query_notion_database(database_id, headers, start_date, end_date, date_property)

    logger.info("📊 Retrieved %s journal entries for date range", len(pages))

    if not pages:
        logger.warning("⚠️ No entries found for this date range")
        return None

    output_sections = []
    week_number = start_date.isocalendar()[1]

    output_sections.append("=" * 80)
    output_sections.append(f"WEEKLY JOURNAL SUMMARY - WEEK # {week_number}")
    output_sections.append(f"Period: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
    output_sections.append("=" * 80)
    output_sections.append("")

    for field in config["fields"]:
        section_title = field.get("section_title", field["name"].upper())
        description = field.get("description", "")
        
        logger.info("🔄 Processing field: %s", section_title)

        field_data = process_field_from_pages(pages, field)

        if field_data:
            output_sections.append("-" * 80)
            output_sections.append(f"{section_title}")
            if description:
                output_sections.append(f"({description})")
            output_sections.append("-" * 80)
            output_sections.append(field_data)
            output_sections.append("")
        else:
            logger.warning("⚠️ No data found for field: %s", section_title)

    if config.get("checkboxes"):
        logger.info("🔄 Processing checkboxes")
        checkbox_data = process_checkboxes_from_pages(pages, config['checkboxes'])
        
        if checkbox_data:
            output_sections.append("-" * 80)
            output_sections.append("DAILY PRACTICES")
            output_sections.append("(Days practiced this week)")
            output_sections.append("-" * 80)
            output_sections.append(checkbox_data)
            output_sections.append("")

    consolidated_output = "\n".join(output_sections)

    output_dir = config["output"]["directory"]
    filename_pattern = config["output"]["filename_pattern"]
    output_filename = filename_pattern.format(
        start_date=start_date.strftime('%Y-%m-%d'),
        end_date=end_date.strftime('%Y-%m-%d')
    )
    
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    output_path = os.path.join(output_dir, output_filename)
    logger.info("📝 Writing consolidated output: %s", output_filename)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(consolidated_output)

    logger.info("✅ Saved: %s", output_path)
    logger.info("📄 Total characters: %s", len(consolidated_output))

    if not _NON_INTERACTIVE:
        _open_folder_in_explorer(output_dir)

    return output_path

if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
    )

    config_path = os.path.join(os.path.dirname(__file__), "journal_automation.json")
    logger.info("📂 Loading configuration...")
    config = load_json_config(config_path)

    logger.info("🔐 Loading environment variables...")
    env_vars = load_env_variables()

    logger.info("🚀 Processing journal entries...")
    output_path = process_journal_consolidated(config, env_vars)

    if output_path:
        logger.info("✅ Journal processing completed!")
        logger.info("📁 Output file: %s", output_path)
        logger.info("💡 Tip: You can now use this consolidated file to prompt an LLM for weekly summary.")
    else:
        logger.warning("⚠️ No output generated (no entries found)")
