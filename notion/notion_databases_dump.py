# requirements: public
import os
import pandas as pd
from notion_client import Client
import datetime
import time

# requirements: custom functions
from utils import read_params_from_txt_file
from utils import get_column_widths
from utils import apply_column_widths

# sources:
# 2023-08-08 > https://chat.openai.com/c/18520141-8212-42e8-97aa-09f5de9713f2

# Function to update the last_download and n_rows columns in the Excel file
def update_last_download_in_excel(database, excel_path, n_rows, output_path):
    # Read the existing Excel file
    df = pd.read_excel(excel_path)

    # Update the last_download column with the current timestamp for the corresponding row
    df.loc[df['id'] == database['id'], 'last_download'] = pd.Timestamp(time.time(), unit='s')

    # Update the n_rows column with the number of processed rows for the corresponding row
    df.loc[df['id'] == database['id'], 'n_rows'] = n_rows

    # Update the save_path column with the saved excel path
    df.loc[df['id'] == database['id'], 'output_path'] = output_path

    # Save the updated DataFrame to the Excel file, overwriting it
    df.to_excel(excel_path, index=False, engine='openpyxl')

def download_database_data(database, output_folder, api_token, excel_path):
    # Authenticate
    notion = Client(auth=api_token)

    # Get the database properties
    properties = notion.databases.retrieve(database["id"]).get("properties")

    # Initialize variables for pagination
    has_more = True
    start_cursor = None
    data = []
    id_list = []  # variable to store the internal IDs of the rows
    rows_processed = 0
    start_time = datetime.datetime.now()

    # Print the database and start
    print(f"Starting to save database '{database['name']}'")

    # Loop through the paginated results to get all rows
    while has_more:
        # Query the database data with a start cursor if provided
        response = notion.databases.query(database["id"], start_cursor=start_cursor)
        results = response.get("results")
        start_cursor = response.get("next_cursor")

        # Transform the data into a DataFrame
        for row in results:
            record = {}
            id_list.append(row['id'])  # Add the internal ID of the row to the id_list
            for key, value in row["properties"].items():
                record[key] = value.get(properties[key]["type"])
            data.append(record)
            rows_processed += 1

            # Print progress every 100 rows
            if rows_processed % 100 == 0:
                elapsed_time = datetime.datetime.now() - start_time
                print(f"{datetime.datetime.now()} - Processed {rows_processed} rows - Time elapsed: {elapsed_time}")

        # Check if there are more pages to be fetched
        has_more = start_cursor is not None

    # Create a DataFrame from the fetched data
    df = pd.DataFrame(data)

    # Add the 'id' column to the DataFrame using the id_list
    df.insert(0, 'id', id_list)

    # Save the DataFrame as an Excel file
    output_path = os.path.join(output_folder, f"database-dump-{database['name']}.xlsx")

    # Check if output_path is a valid file before processing
    if os.path.isfile(output_path):
        # Before processing the Excel file reads the column widths
        column_widths = get_column_widths(output_path)
    else:
        print(f"'{output_path}' is not a valid file. Skipping reading column widths")
        column_widths = False

    df.to_excel(output_path, index=False, engine='openpyxl')

    # Check if column_widths is valid (in this case, non-empty) before applying
    if column_widths == False:
        print("Column widths are not valid, skipping the step.")
    else:
        # After processing the Excel file recovers the column widths
        apply_column_widths(excel_path, column_widths)


    # Print the saved file information
    print(f"Database '{database['name']}' saved to {output_path} with {len(df)} rows")

    # Update the last_download and n_rows columns in the Excel file for the current database
    update_last_download_in_excel(database, excel_path, len(df),output_path)

    return output_path

def read_database_list(excel_path):
    df = pd.read_excel(excel_path)
    return df[df["ind_download"] == 1]

# Main execution

# Load the parameters from the text file
params_file_path = r"C:\Mis Datos en Local\temporal\python\notion-params.txt"
params = read_params_from_txt_file(params_file_path)

# Get the api_token
api_token = params['api_token']

# Set the Excel file path and the dump path
excel_path = params['excel_path']
dump_path = params['dump_path']

# Read the database list from the Excel file and filter the rows with "download = 1"
databases_to_download = read_database_list(excel_path)

# Download the data for each selected database
for _, database in databases_to_download.iterrows():
    output_path = download_database_data(database, dump_path, api_token, excel_path)





