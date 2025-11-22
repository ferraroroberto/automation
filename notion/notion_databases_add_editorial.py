# source chatGPT > https://chatgpt.com/c/6724bbf2-a618-8009-a4e3-a1f6e6828ae0

import pandas as pd
from notion_client import Client
from datetime import datetime, timedelta
from utils import read_params_from_txt_file

# Initialize Notion Client using parameters file
def init_notion_client(params_file_path):
    params = read_params_from_txt_file(params_file_path)
    api_token = params['api_token']
    return Client(auth=api_token)

# Format database ID with hyphens
def format_database_id(database_id):
    if len(database_id) == 32:
        # Insert hyphens to convert into UUID format
        return f"{database_id[:8]}-{database_id[8:12]}-{database_id[12:16]}-{database_id[16:20]}-{database_id[20:]}"
    return database_id

# Check existing records and add missing dates
def add_missing_dates(notion, database_id, start_date, end_date):
    # Format for Notion date properties
    date_format = "%Y-%m-%d"
    day_text_format = "%Y%m%d"
    day_of_week_format = "%a"
    
    # Convert input dates to datetime objects
    start = datetime.strptime(start_date, date_format)
    end = datetime.strptime(end_date, date_format)

    # Retrieve all existing dates from the database
    existing_dates = get_existing_dates(notion, database_id)

    # Debug: Confirm database existence
    if existing_dates is None:
        print("Error: Could not retrieve existing dates. Please check database ID.")
        return

    # Debug: Show how many existing dates were found
    print(f"Found {len(existing_dates)} existing date(s) in database")
    if existing_dates:
        sample_dates = sorted(list(existing_dates))[:5]  # Show first 5 dates
        print(f"Sample existing dates: {sample_dates}")
    print()  # Blank line after sample dates

    # Generate missing date records
    new_entries = []
    current_date = start

    while current_date <= end:
        date_str = current_date.strftime(date_format)  # YYYY-MM-DD format for ISO 8601
        if date_str not in existing_dates:
            # Create data fields
            day_text = current_date.strftime(day_text_format)  # YYYYMMDD as text
            day_of_week = current_date.strftime(day_of_week_format).lower()

            # Add to Notion
            new_entry = add_date_record(notion, database_id, day_text, date_str, day_of_week)
            if new_entry:
                print(f"Added record: day = {day_text}, date = {date_str}, DoW = {day_of_week}, Record ID = {new_entry['id']}")
                new_entries.append(new_entry)
        else:
            # Debug: Show which dates are being skipped
            print(f"Skipping existing date: {date_str}")
        
        current_date += timedelta(days=1)

    # Summary
    print(f"\nTotal new records added: {len(new_entries)}\n")  # Line break after total records added
    for entry in new_entries:
        print(f"Record ID: {entry['id']}, day: {entry['properties']['day']['title'][0]['plain_text']}")

# Helper: Retrieve existing dates
def get_existing_dates(notion, database_id):
    try:
        # Format database ID with hyphens if needed
        formatted_db_id = format_database_id(database_id)
        
        # Get database details to retrieve data source ID
        database_details = notion.databases.retrieve(formatted_db_id)
        data_sources = database_details.get("data_sources") or []
        if not data_sources:
            print("Error: No data sources associated with database")
            return None
        data_source_id = data_sources[0]["id"]
        
        # Query database using data source with pagination
        print("Loading database records...")
        existing_dates = set()
        has_more = True
        start_cursor = None
        cumulative_count = 0
        
        while has_more:
            response = notion.data_sources.query(
                data_source_id,
                start_cursor=start_cursor,
            )
            
            # Check if response is valid
            if response is None or "results" not in response:
                print("Error: No results returned. Check database permissions and structure.")
                return None
            
            results = response.get("results", [])
            start_cursor = response.get("next_cursor")
            
            # Log progress for each API call
            records_in_batch = len(results)
            cumulative_count += records_in_batch
            print(f"Fetched {records_in_batch} records (cumulative: {cumulative_count})")
            
            # Collect existing dates from the "date" property
            for item in results:
                # Check if "date" property exists and has a "start" field
                date_property = item["properties"].get("date", {}).get("date")
                if date_property and date_property.get("start"):
                    date_value = date_property["start"]
                    # Normalize date to YYYY-MM-DD format (strip time if present)
                    if "T" in date_value:
                        date_value = date_value.split("T")[0]
                    existing_dates.add(date_value)
            
            # Check if there are more pages to be fetched
            has_more = start_cursor is not None
        
        return existing_dates
    
    except Exception as e:
        print(f"Error retrieving existing dates: {e}")
        return None

# Helper: Add a new record to the Notion database
def add_date_record(notion, database_id, day, date, day_of_week):
    try:
        # Format database ID with hyphens if needed
        formatted_db_id = format_database_id(database_id)
        return notion.pages.create(
            parent={"database_id": formatted_db_id},
            properties={
                "day": {"title": [{"text": {"content": day}}]},
                "date": {"date": {"start": date}},
                "DoW": {"rich_text": [{"text": {"content": day_of_week}}]}
            }
        )
    except Exception as e:
        print(f"Error adding record for {date}: {e}")
        return None

# Main execution
if __name__ == "__main__":
    # Specify the path to the parameters file
    params_file_path = r"C:\Mis Datos en Local\temporal\python\notion-params.txt"
    database_id = "ee23dec3222e4412baaf0d129d0bda9c"

    notion = init_notion_client(params_file_path)

    # Ask for date range in YYYYMMDD format
    start_date_input = input("Enter the start date (YYYYMMDD): ")
    end_date_input = input("Enter the end date (YYYYMMDD): ")
    print()  # Line break after start date input

    # Convert input dates to required format
    start_date = datetime.strptime(start_date_input, "%Y%m%d").strftime("%Y-%m-%d")
    end_date = datetime.strptime(end_date_input, "%Y%m%d").strftime("%Y-%m-%d")

    # Run the function to add missing dates
    add_missing_dates(notion, database_id, start_date, end_date)
