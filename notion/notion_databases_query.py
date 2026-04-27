# requirements: public
import logging
import pandas as pd
from notion_client import Client

# requirements: custom functions
from utils import read_params_from_txt_file

log = logging.getLogger(__name__)

def get_database_list():
    databases = notion.search(filter={"property": "object", "value": "database"}).get("results", [])
    database_list = []
    for db in databases:
        # Remove "-" characters from the database ID
        clean_id = db["id"].replace("-", "")
        # Check if the database has a title
        if db["title"]:
            database_list.append({"name": db["title"][0]["plain_text"], "id": clean_id, "url": f"{workspace_url}{clean_id}?v=b"})
        else:
            # Assign a default name if the title is missing
            database_list.append({"name": "Untitled", "id": clean_id, "url": f"{workspace_url}{clean_id}?v=b"})
    return database_list

def save_to_excel(databases, db_excel_path):
    df = pd.DataFrame(databases, columns=["name", "id", "url"])
    df.to_excel(db_excel_path, index=False, engine='openpyxl')

# Main execution

# Load the parameters from the text file
params_file_path = r"C:\Mis Datos en Local\temporal\python\notion-params.txt"
params = read_params_from_txt_file(params_file_path)

# Get the api_token
api_token = params['api_token']

# Authenticate
notion = Client(auth=api_token)

# Get the Excel file path and the Notion workspace URL
db_excel_path = params['db_excel_path']
workspace_url = params['workspace_url']

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    databases = get_database_list()
    save_to_excel(databases, db_excel_path)
    log.info("Databases saved to %s", db_excel_path)
