import pandas as pd
from notion_client import Client, APIResponseError
import time
from utils import read_params_from_txt_file

# chatGPT source > https://chat.openai.com/c/4db5a8e0-1344-4fb7-8952-d96607a88b9e
# chatGPT source > https://chat.openai.com/c/2bfb9ae6-233d-46fe-8171-69e7e1929866
# chatGPT source > https://chat.openai.com/c/025bb803-a79e-4be1-851f-fbb249db6e38

params_file_path = r"C:\Mis Datos en Local\temporal\python\notion-params.txt"
params = read_params_from_txt_file(params_file_path)

sync_path = params['sync_path']
api_token = params['api_token']
verbose = params['verbose']

print("Loading Excel file...")
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

print("Initializing Notion client...")
client = Client(auth=api_token)

# Step 1: Fetch Notion database details
try:
    database_details = client.databases.retrieve(notion_db_id)
    database_name = database_details['title'][0]['text']['content']

    # Now, query the database to get the number of rows (pages)
    total_rows = len(client.databases.query(notion_db_id)['results'])

    print(f"Database Name: {database_name}")
    print(f"Total Rows in Notion: {total_rows}")

except APIResponseError as e:
    print(f"Failed to fetch database details: {e}")

# Step 2: Ask for confirmation to proceed
input("Press Enter to continue...")

# Initialize counters
archived_pages_counter = 0
created_pages_counter = 0

# If clean flag is True, archive all pages in the database
if clean_db:
    print("Archiving all pages in the Notion database...")
    try:
        response = client.databases.query(notion_db_id)
        for page in response['results']:
            client.pages.update(page['id'], properties={"archived": {"checkbox": True}})
            archived_pages_counter += 1
            print(f"Archiving page with ID: {page['id']}. Total archived pages: {archived_pages_counter}")
            time.sleep(0.0)  # Adjust delay as needed
    except APIResponseError as e:
        print(f"Failed to archive pages: {e}")

print("Creating all pages in the Excel database...")
for index, row in df.iterrows():
    time.sleep(0.0)  # Adjust delay as needed

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
        if verbose == True: print(f"Creating new page with data: {page_data}")
        new_page = client.pages.create(**page_data)
        created_pages_counter += 1
        print(f"Created new page with ID: {new_page['id']}. Total created pages: {created_pages_counter}")
    except APIResponseError as e:
        print(f"Failed to create new page: {e}")
