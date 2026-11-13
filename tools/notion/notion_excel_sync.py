import logging

import pandas as pd
from notion_client import APIResponseError, Client
from notion_client.helpers import collect_paginated_api
from utils import (
    DEFAULT_PARAMS_FILE,
    read_params_from_txt_file,
    require_notion_api_token,
)

log = logging.getLogger(__name__)


def _list_page_ids(client: Client, data_source_id: str) -> list:
    """Collect every page id of a data source, paginating fully."""
    # Gather ids before archiving: archiving mutates the query result set, so
    # paging while archiving would skip pages that shift out of later pages.
    pages = collect_paginated_api(client.data_sources.query, data_source_id=data_source_id)
    return [page['id'] for page in pages]


def _build_page_data(row: pd.Series, columns, metadata_dict: dict, notion_db_id: str) -> dict:
    """Map one Excel row to the Notion page payload using the metadata sheet."""
    page_data = {"parent": {"database_id": notion_db_id}, "properties": {}}

    for col in columns:
        # Skip column if it doesn't exist in metadata_dict or if 'keep' value is False
        if col not in metadata_dict or not metadata_dict[col]['keep'] or pd.isnull(row[col]):
            continue

        notion_col = metadata_dict[col]['notionColumn']
        datatype = metadata_dict[col]['type']

        if datatype == 'title':
            page_data["properties"][notion_col] = {"title": [{"text": {"content": str(row[col])}}]}
        elif datatype == 'date':
            page_data["properties"][notion_col] = {"date": {"start": row[col].strftime('%Y-%m-%d')}}
        elif datatype == 'number':
            page_data["properties"][notion_col] = {"number": row[col]}
        elif datatype == 'url':
            page_data["properties"][notion_col] = {"url": str(row[col])}
        elif datatype == 'relation':
            related_page_ids = str(row[col]).split(',')
            page_data["properties"][notion_col] = {"relation": [{"id": page_id.strip()} for page_id in related_page_ids]}
        elif datatype == 'rich_text':
            page_data["properties"][notion_col] = {"rich_text": [{"text": {"content": str(row[col])}}]}
        elif datatype == 'checkbox':
            page_data["properties"][notion_col] = {"checkbox": bool(row[col])}
    return page_data


def main():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    params_file_path = DEFAULT_PARAMS_FILE
    params = read_params_from_txt_file(params_file_path)

    sync_path = params['sync_path']
    api_token = require_notion_api_token()
    verbose = params['verbose'] == 'True'

    log.info("Loading Excel file...")
    xlsx = pd.ExcelFile(sync_path)

    # Load the data, metadata and db sheets
    df = pd.read_excel(xlsx, 'data')
    metadata_df = pd.read_excel(xlsx, 'metadata')
    db_df = pd.read_excel(xlsx, 'db')

    # Extract the Notion database ID
    notion_db_id = db_df.loc[0, 'id']

    # Extract the clean flag
    clean_db = db_df.loc[0, 'clean']

    # Generate a dictionary from the metadata
    metadata_dict = metadata_df.set_index('excelColumn')[['notionColumn', 'type', 'keep']].T.to_dict()

    log.info("Initializing Notion client...")
    client = Client(auth=api_token)

    # Step 1: Fetch Notion database details
    try:
        database_details = client.databases.retrieve(notion_db_id)
        database_name = database_details['title'][0]['text']['content']

        # notion-client 2.x removed Client.databases.query; rows are queried
        # through the database's data source instead.
        data_sources = database_details.get("data_sources") or []
        if not data_sources:
            log.error("No data sources on this database (Notion API 2025+). Check integration access.")
            return
        data_source_id = data_sources[0]["id"]

        total_rows = len(_list_page_ids(client, data_source_id))

        log.info("Database Name: %s", database_name)
        log.info("Total Rows in Notion: %d", total_rows)

    except APIResponseError as e:
        log.error("Failed to fetch database details: %s", e)
        return

    # Step 2: Ask for confirmation to proceed
    input("Press Enter to continue...")

    # Initialize counters
    archived_pages_counter = 0
    created_pages_counter = 0

    # If clean flag is True, archive all pages in the database
    if clean_db:
        log.info("Archiving all pages in the Notion database...")
        try:
            # Re-list after the confirmation prompt so rows added while it waited are archived too.
            for page_id in _list_page_ids(client, data_source_id):
                # archived is a top-level page field in the Notion API, not a property
                client.pages.update(page_id, archived=True)
                archived_pages_counter += 1
                log.info("Archiving page with ID: %s. Total archived pages: %d", page_id, archived_pages_counter)
        except APIResponseError as e:
            log.error("Failed to archive pages: %s", e)

    log.info("Creating all pages in the Excel database...")
    for index, row in df.iterrows():
        page_data = _build_page_data(row, df.columns, metadata_dict, notion_db_id)

        try:
            if verbose == True: log.debug("Creating new page with data: %s", page_data)
            new_page = client.pages.create(**page_data)
            created_pages_counter += 1
            log.info("Created new page with ID: %s. Total created pages: %d", new_page['id'], created_pages_counter)
        except APIResponseError as e:
            log.error("Failed to create new page: %s", e)


if __name__ == "__main__":
    main()
