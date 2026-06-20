# Run with repo venv: ..\.venv\Scripts\python notion_databases_add_editorial.py (from this folder), or notion_databases_add_editorial.bat



import argparse

import logging

import sys

from calendar import monthrange

from datetime import datetime, timedelta
from typing import Optional

from notion_client import Client



from utils import read_params_from_txt_file, DEFAULT_PARAMS_FILE





def setup_logging(debug: bool = False):

    """Set up logging configuration (same style as normalize_url.py)."""

    level = logging.DEBUG if debug else logging.INFO

    logging.basicConfig(

        level=level,

        format="%(asctime)s - %(levelname)s - %(message)s",

        handlers=[logging.StreamHandler(sys.stdout)],

    )





def editorial_date_range(now: Optional[datetime] = None) -> tuple[str, str]:

    """

    First day of the current calendar month through last day of the next calendar month (YYYY-MM-DD).

    """

    today = now or datetime.now()

    start = datetime(today.year, today.month, 1)

    if today.month == 12:

        end_year, end_month = today.year + 1, 1

    else:

        end_year, end_month = today.year, today.month + 1

    last_day = monthrange(end_year, end_month)[1]

    end = datetime(end_year, end_month, last_day)

    fmt = "%Y-%m-%d"

    return start.strftime(fmt), end.strftime(fmt)





# Initialize Notion Client using parameters file

def init_notion_client(params_file_path, timeout_ms=180_000):

    params = read_params_from_txt_file(params_file_path)

    api_token = params["api_token"]

    # Default SDK timeout is 60s; large DBs or slow links often need more.

    return Client(auth=api_token, timeout_ms=timeout_ms)





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



    logging.info(f"📅 Range: {start_date} → {end_date}")



    # Retrieve existing dates in range (filtered query = fewer pages, faster than full scan)

    existing_dates = get_existing_dates(notion, database_id, start_date, end_date)



    if existing_dates is None:

        logging.error("Could not retrieve existing dates. Check database ID and integration access.")

        return



    logging.info(f"📊 Found {len(existing_dates)} existing date(s) in database")

    if existing_dates:

        sample_dates = sorted(list(existing_dates))[:5]

        logging.info(f"Sample existing dates: {sample_dates}")



    # Generate missing date records

    new_entries = []

    current_date = start



    while current_date <= end:

        date_str = current_date.strftime(date_format)  # YYYY-MM-DD format for ISO 8601

        if date_str not in existing_dates:

            day_text = current_date.strftime(day_text_format)  # YYYYMMDD as text

            day_of_week = current_date.strftime(day_of_week_format).lower()



            new_entry = add_date_record(notion, database_id, day_text, date_str, day_of_week)

            if new_entry:

                logging.info(

                    f"📝 Added: day={day_text}, date={date_str}, DoW={day_of_week}, id={new_entry['id']}"

                )

                new_entries.append(new_entry)

        else:

            logging.debug(f"Skipping existing date: {date_str}")



        current_date += timedelta(days=1)



    logging.info(f"✅ Total new records added: {len(new_entries)}")

    for entry in new_entries:

        title = entry["properties"]["day"]["title"][0]["plain_text"]

        logging.debug(f"Record ID: {entry['id']}, day: {title}")





# Helper: Retrieve existing dates in [range_start, range_end] (YYYY-MM-DD)

def get_existing_dates(notion, database_id, range_start, range_end):

    try:

        formatted_db_id = format_database_id(database_id)



        # notion-client 2.x removed Client.databases.query; use data_sources.query instead.

        database_details = notion.databases.retrieve(formatted_db_id)

        data_sources = database_details.get("data_sources") or []

        if not data_sources:

            logging.error(

                "No data sources on this database (Notion API 2025+). Check integration access."

            )

            return None

        data_source_id = data_sources[0]["id"]



        date_filter = {

            "and": [

                {"property": "date", "date": {"on_or_after": range_start}},

                {"property": "date", "date": {"on_or_before": range_end}},

            ]

        }



        logging.info("Loading database records...")

        existing_dates = set()

        has_more = True

        start_cursor = None

        cumulative_count = 0



        while has_more:

            response = notion.data_sources.query(

                data_source_id,

                filter=date_filter,

                page_size=100,

                start_cursor=start_cursor,

            )



            if response is None or "results" not in response:

                logging.error("No results returned. Check database permissions and structure.")

                return None



            results = response.get("results", [])

            start_cursor = response.get("next_cursor")



            records_in_batch = len(results)

            cumulative_count += records_in_batch

            logging.info(f"📥 Fetched {records_in_batch} records (cumulative: {cumulative_count})")



            for item in results:

                date_property = item["properties"].get("date", {}).get("date")

                if date_property and date_property.get("start"):

                    date_value = date_property["start"]

                    if "T" in date_value:

                        date_value = date_value.split("T")[0]

                    existing_dates.add(date_value)



            has_more = start_cursor is not None



        return existing_dates



    except Exception as e:

        logging.error(f"Error retrieving existing dates: {e}")

        return None





# Helper: Add a new record to the Notion database

def add_date_record(notion, database_id, day, date, day_of_week):

    try:

        formatted_db_id = format_database_id(database_id)

        return notion.pages.create(

            parent={"database_id": formatted_db_id},

            properties={

                "day": {"title": [{"text": {"content": day}}]},

                "date": {"date": {"start": date}},

                "DoW": {"rich_text": [{"text": {"content": day_of_week}}]},

            },

        )

    except Exception as e:

        logging.error(f"Error adding record for {date}: {e}")

        return None





def main():

    parser = argparse.ArgumentParser(

        description="Add missing editorial calendar rows for the current and next calendar month."

    )

    parser.add_argument(

        "--params",

        type=str,

        default=DEFAULT_PARAMS_FILE,

        help="Path to notion-params.txt (api_token)",

    )

    parser.add_argument(

        "--database-id",

        type=str,

        default="ee23dec3222e4412baaf0d129d0bda9c",

        help="Notion database ID (32-char or UUID)",

    )

    parser.add_argument("--debug", action="store_true", help="Enable debug logging")



    args = parser.parse_args()

    setup_logging(args.debug)



    start_date, end_date = editorial_date_range()

    logging.info("✅ Editorial DB sync starting (current month + next month)")



    try:

        notion = init_notion_client(args.params)

        add_missing_dates(notion, args.database_id, start_date, end_date)

    except Exception as e:

        logging.error(f"❌ Fatal error: {e}")

        if args.debug:

            logging.exception("Full traceback:")

        sys.exit(1)





if __name__ == "__main__":

    main()

