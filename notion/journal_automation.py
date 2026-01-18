# chatGPT source 2023-05-01 > https://chat.openai.com/c/de2353d2-d8c8-41c0-b9d2-8317c39c8c2c
# chatGPT source 2023-09-26 > https://chat.openai.com/c/2f7bf670-4b1b-4cd1-874d-0bb54174bddc
# chatGPT source 2023-10-08 > https://chat.openai.com/c/026c7ef9-a6d2-4304-ac37-452d1896c592 

# requirements: public
import os
import requests
from datetime import datetime, timedelta

# requirements: custom functions
from utils import load_json_config, load_env_variables

# Global variable to store the adjusted date
adjusted_date = None

def get_adjusted_date(timedelta_value=None):
    """Get the adjusted date for processing."""
    global adjusted_date
    
    if adjusted_date is not None:
        return adjusted_date

    # Get the current date
    current_date = datetime.now()

    # Ask the user if they want to apply a timedelta to the current date if not provided
    if timedelta_value is None:
        timedelta_value = input("📅 Enter an integer number to apply a timedelta to the current date (default is 0): ")
        try:
            timedelta_value = int(timedelta_value)
        except ValueError:
            timedelta_value = 0

    # Apply the timedelta to the current date
    adjusted_date = current_date + timedelta(days=timedelta_value)
    
    return adjusted_date

def calculate_week_range(current_date):
    """
    Calculate the date range for the previous week (Monday to Sunday).
    If today is Sunday, use today as the end date.
    
    Returns:
        tuple: (start_date, end_date) as datetime objects
    """
    # Check if today is Sunday (weekday() returns 6 for Sunday)
    if current_date.weekday() == 6:  # Sunday
        end_date = current_date
    else:
        # Calculate the previous Sunday
        days_since_sunday = (current_date.weekday() + 1) % 7
        if days_since_sunday == 0:
            days_since_sunday = 7
        end_date = current_date - timedelta(days=days_since_sunday)
    
    # Calculate Monday (6 days before Sunday)
    start_date = end_date - timedelta(days=6)
    
    # Set to start and end of day
    start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)
    
    return start_date, end_date

def query_notion_database(database_id, headers, start_date, end_date, date_property):
    """
    Query Notion database for journal entries in the date range.
    
    Args:
        database_id: Notion database ID
        headers: API headers with authorization
        start_date: Start date for filtering
        end_date: End date for filtering
        date_property: Name of the date property in Notion
        
    Returns:
        List of page objects from Notion API
    """
    print(f"🔍 Querying Notion database...")
    
    # Convert dates to ISO format for Notion API
    start_iso = start_date.isoformat()
    end_iso = end_date.isoformat()
    
    filter_body = {
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
    
    pages = []
    start_cursor = None
    
    while True:
        if start_cursor:
            filter_body["start_cursor"] = start_cursor
        
        try:
            response = requests.post(
                f"https://api.notion.com/v1/databases/{database_id}/query",
                headers=headers,
                json=filter_body
            )
            response.raise_for_status()
            
            data = response.json()
            results = data.get('results', [])
            
            if not results:
                break
            
            pages.extend(results)
            print(f"📥 Retrieved {len(results)} entries (total: {len(pages)})")
            
            if not data.get('has_more', False):
                break
            
            start_cursor = data.get('next_cursor')
            
        except requests.exceptions.RequestException as e:
            print(f"❌ Notion API error: {e}")
            if hasattr(e, 'response') and e.response:
                print(f"📊 Status: {e.response.status_code}")
                print(f"Response: {e.response.text[:500]}")
            raise
    
    return pages

def extract_property_value(properties, property_name):
    """
    Extract value from a Notion property.
    
    Args:
        properties: Properties dict from Notion page
        property_name: Name of the property to extract
        
    Returns:
        String value or empty string if not found
    """
    prop = properties.get(property_name, {})
    prop_type = prop.get('type')
    
    if not prop_type:
        return ""
    
    # Handle different property types
    if prop_type == 'title':
        title_content = prop.get('title', [])
        return ''.join([segment.get('plain_text', '') for segment in title_content]).strip()
    
    elif prop_type == 'rich_text':
        rich_text_content = prop.get('rich_text', [])
        return ''.join([segment.get('plain_text', '') for segment in rich_text_content]).strip()
    
    elif prop_type == 'multi_select':
        options = prop.get('multi_select', [])
        return ', '.join([opt.get('name', '') for opt in options])
    
    elif prop_type == 'select':
        option = prop.get('select', {})
        return option.get('name', '') if option else ""
    
    elif prop_type == 'checkbox':
        return "✓" if prop.get('checkbox', False) else ""
    
    elif prop_type == 'date':
        date_obj = prop.get('date', {})
        if date_obj:
            return date_obj.get('start', '')
        return ""
    
    return ""

def process_field_from_pages(pages, field_config):
    """
    Process a single field from the list of Notion pages.
    
    Args:
        pages: List of Notion page objects
        field_config: Configuration dictionary for this field
        
    Returns:
        String containing processed data for this field with frequency counts
    """
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

def process_checkboxes_from_pages(pages, checkbox_configs):
    """
    Process checkbox fields and count how many days they were checked.
    
    Args:
        pages: List of Notion page objects
        checkbox_configs: List of checkbox configuration dictionaries
        
    Returns:
        String containing checkbox summary (e.g., "self-centered: 5 days out of 7")
    """
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

def process_journal_consolidated(config, env_vars):
    """
    Process journal entries from Notion API and create consolidated output.
    
    Args:
        config: Configuration dictionary loaded from JSON file
        env_vars: Environment variables dictionary
    """
    
    # Get the adjusted date
    current_date = get_adjusted_date()
    
    # Calculate date range
    start_date, end_date = calculate_week_range(current_date)
    
    print(f"📅 Date range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
    
    # Setup Notion API credentials
    notion_api_key = env_vars.get('notion_api_token')
    if not notion_api_key:
        raise ValueError("❌ NOTION_API_TOKEN not found in environment variables")
    
    headers = {
        'Authorization': f'Bearer {notion_api_key}',
        'Notion-Version': '2022-06-28',
        'Content-Type': 'application/json'
    }
    
    # Get database configuration
    database_id = config['input']['database_id']
    date_property = config['input'].get('date_property', 'Date')
    
    # Query Notion database
    pages = query_notion_database(database_id, headers, start_date, end_date, date_property)
    
    print(f"📊 Retrieved {len(pages)} journal entries for date range")
    
    if len(pages) == 0:
        print("⚠️  No entries found for this date range")
        return None
    
    # Build consolidated output
    output_sections = []
    
    # Calculate week number (ISO week number)
    week_number = start_date.isocalendar()[1]
    
    # Add header
    output_sections.append("=" * 80)
    output_sections.append(f"WEEKLY JOURNAL SUMMARY - WEEK # {week_number}")
    output_sections.append(f"Period: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
    output_sections.append("=" * 80)
    output_sections.append("")
    
    # Process each field
    for field in config['fields']:
        section_title = field.get('section_title', field['name'].upper())
        description = field.get('description', '')
        
        print(f"🔄 Processing field: {section_title}")
        
        field_data = process_field_from_pages(pages, field)
        
        if field_data:  # Only add section if there's data
            output_sections.append("-" * 80)
            output_sections.append(f"{section_title}")
            if description:
                output_sections.append(f"({description})")
            output_sections.append("-" * 80)
            output_sections.append(field_data)
            output_sections.append("")
        else:
            print(f"⚠️  No data found for field: {section_title}")
    
    # Process checkboxes if configured
    if 'checkboxes' in config and config['checkboxes']:
        print(f"🔄 Processing checkboxes")
        checkbox_data = process_checkboxes_from_pages(pages, config['checkboxes'])
        
        if checkbox_data:
            output_sections.append("-" * 80)
            output_sections.append("DAILY PRACTICES")
            output_sections.append("(Days practiced this week)")
            output_sections.append("-" * 80)
            output_sections.append(checkbox_data)
            output_sections.append("")
    
    # Combine all sections
    consolidated_output = "\n".join(output_sections)
    
    # Create output filename
    output_dir = config['output']['directory']
    filename_pattern = config['output']['filename_pattern']
    output_filename = filename_pattern.format(
        start_date=start_date.strftime('%Y-%m-%d'),
        end_date=end_date.strftime('%Y-%m-%d')
    )
    
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    # Save to file
    output_path = os.path.join(output_dir, output_filename)
    print(f"📝 Writing consolidated output: {output_filename}")
    
    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(consolidated_output)
    
    print(f"✅ Saved: {output_path}")
    print(f"📄 Total characters: {len(consolidated_output)}")
    
    return output_path

# Main execution
if __name__ == "__main__":
    # Load configuration
    config_path = os.path.join(os.path.dirname(__file__), "journal_automation.json")
    print("📂 Loading configuration...")
    config = load_json_config(config_path)
    
    # Load environment variables
    print("🔐 Loading environment variables...")
    env_vars = load_env_variables()
    
    print("\n🚀 Processing journal entries...")
    output_path = process_journal_consolidated(config, env_vars)
    
    if output_path:
        print("\n✅ Journal processing completed!")
        print(f"📁 Output file: {output_path}")
        print("\n💡 Tip: You can now use this consolidated file to prompt an LLM for weekly summary analysis.")
    else:
        print("\n⚠️  No output generated (no entries found)")
