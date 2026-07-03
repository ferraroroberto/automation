import logging
import pandas as pd
from notion_client import Client, APIResponseError
from utils import read_params_from_txt_file, DEFAULT_PARAMS_FILE

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
log = logging.getLogger(__name__)

# chatGPT source > https://chat.openai.com/c/4db5a8e0-1344-4fb7-8952-d96607a88b9e
# chatGPT source > https://chat.openai.com/c/2bfb9ae6-233d-46fe-8171-69e7e1929866
# chatGPT source > https://chat.openai.com/c/025bb803-a79e-4be1-851f-fbb249db6e38

params_file_path = DEFAULT_PARAMS_FILE
params = read_params_from_txt_file(params_file_path)

sync_path = params['sync_path']
api_token = params['api_token']
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

    # Now, query the database to get the number of rows (pages), paginating fully
    total_rows = 0
    start_cursor = None
    while True:
        page = client.databases.query(notion_db_id, start_cursor=start_cursor)
        total_rows += len(page['results'])
        if not page.get('has_more'):
            break
        start_cursor = page.get('next_cursor')

    log.info("Database Name: %s", database_name)
    log.info("Total Rows in Notion: %d", total_rows)

except APIResponseError as e:
    log.error("Failed to fetch database details: %s", e)

# Step 2: Ask for confirmation to proceed
input("Press Enter to continue...")

# Initialize counters
archived_pages_counter = 0
created_pages_counter = 0

# If clean flag is True, archive all pages in the database
if clean_db:
    log.info("Archiving all pages in the Notion database...")
    try:
        # Collect every page id first (paginating via next_cursor); archiving
        # mutates the query result set, so gather ids before archiving to avoid
        # skipping pages that shift out of later pages.
        page_ids = []
        start_cursor = None
        while True:
            response = client.databases.query(notion_db_id, start_cursor=start_cursor)
            page_ids.extend(page['id'] for page in response['results'])
            if not response.get('has_more'):
                break
            start_cursor = response.get('next_cursor')

        for page_id in page_ids:
            # archived is a top-level page field in the Notion API, not a property
            client.pages.update(page_id, archived=True)
            archived_pages_counter += 1
            log.info("Archiving page with ID: %s. Total archived pages: %d", page_id, archived_pages_counter)
    except APIResponseError as e:
        log.error("Failed to archive pages: %s", e)

log.info("Creating all pages in the Excel database...")
for index, row in df.iterrows():
    page_data = {"parent": {"database_id": notion_db_id}, "properties": {}}

    for col in df.columns:
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

    try:
        if verbose == True: log.debug("Creating new page with data: %s", page_data)
        new_page = client.pages.create(**page_data)
        created_pages_counter += 1
        log.info("Created new page with ID: %s. Total created pages: %d", new_page['id'], created_pages_counter)
    except APIResponseError as e:
        log.error("Failed to create new page: %s", e)
