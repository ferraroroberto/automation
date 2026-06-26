import logging
import os
import pandas as pd
import numpy as np
from datetime import datetime

log = logging.getLogger(__name__)

# Source folder where the CSV files are located — set the TODOIST_SOURCE_FOLDER environment variable on each machine
HARDCODED_SOURCE_FOLDER = os.getenv("TODOIST_SOURCE_FOLDER", "")


# Function to determine if DATE is an actual date or a recurring pattern
def is_exact_date(date_str):
    current_year = datetime.now().year
    try:
        # Try parsing the date normally
        date = pd.to_datetime(date_str)
        return 1, date.strftime("%Y-%m-%d"), date
    except (ValueError, TypeError):
        # Check if the date is of the format '31 Aug', missing a year
        try:
            date_with_year = pd.to_datetime(f"{date_str} {current_year}")
            return 1, date_with_year.strftime("%Y-%m-%d"), date_with_year
        except (ValueError, TypeError):
            return 0, "", None


def determine_past_future(exact_date):
    if exact_date is None:
        return -1  # Not a valid date
    today = datetime.now().date()
    if exact_date.date() < today:
        return 1  # Date is in the past
    elif exact_date.date() > today:
        return 0  # Date is in the future
    else:
        return 0  # Date is today or in the future (considered as future)


# Function to split task content into TITLE and URL
def extract_title_url(content):
    if "[" in content and "]" in content and "(" in content and ")" in content:
        title = content.split("[")[1].split("]")[0]
        url = content.split("(")[1].split(")")[0]
        return title, url
    return content, ""


def process_file(filepath, df_list, task_id_counter, note_id_counter, verbose):
    # Extract the project name and ID from the filename
    filename = os.path.basename(filepath)
    project_name = filename.split(' [')[0]
    project_id = filename.split('[')[-1].split(']')[0]

    # Load the file
    df = pd.read_csv(filepath)

    # Check if file is empty
    if df.empty:
        log.debug("Skipping empty file: %s", filepath)
        return task_id_counter, note_id_counter

    log.info("Processing file: %s", filepath)

    last_task_id = None

    for idx, row in df.iterrows():
        # Handle NaN values in 'TYPE'
        row_type = str(row['TYPE']).strip().lower() if pd.notna(row['TYPE']) else ''

        # Concatenate with previous row if the type is missing
        if not row_type:
            df_list[-1]['COMMENT'] += " " + str(row['CONTENT'])
            continue

        # Processing task rows
        if row_type == 'task':
            # Generate a new unique ID for the task with zero-padded format
            task_id = f"TASK_{task_id_counter:04d}"
            task_id_counter += 1
            last_task_id = task_id

            # Extract TITLE and URL from CONTENT
            title, url = extract_title_url(str(row['CONTENT']))

            # Determine if DATE is exact or not, and calculate IND_PAST
            ind_date, exact_date_str, exact_date = is_exact_date(row['DATE'])
            ind_past = determine_past_future(exact_date)

            # Append task to the list
            df_list.append({
                'PROJECT': project_name,
                'PROJECT_ID': project_id,
                'ID': task_id,
                'TYPE': 'task',
                'TITLE': title,
                'URL': url,
                'DESCRIPTION': row['DESCRIPTION'],
                'PRIORITY': row['PRIORITY'],
                'INDENT': row['INDENT'],
                'AUTHOR': row['AUTHOR'],
                'RESPONSIBLE': row['RESPONSIBLE'],
                'DATE': row['DATE'],
                'IND_DATE': ind_date,
                'EXACT_DATE': exact_date_str,
                'IND_PAST': ind_past,
                'DATE_LANG': row['DATE_LANG'],
                'TIMEZONE': row['TIMEZONE'],
                'DURATION': row['DURATION'],
                'DURATION_UNIT': row['DURATION_UNIT'],
                'COMMENT': ''
            })

            log.debug("Processed TASK: %s, TITLE: %s", task_id, title)

        # Processing note rows
        elif row_type == 'note':
            # Generate a new unique ID for the note with zero-padded format
            note_id = f"NOTE_{note_id_counter:04d}"
            note_id_counter += 1

            # Determine if DATE is exact or not, and calculate IND_PAST
            ind_date, exact_date_str, exact_date = is_exact_date(row['DATE'])
            ind_past = determine_past_future(exact_date)

            # Append note to the list
            df_list.append({
                'PROJECT': project_name,
                'PROJECT_ID': project_id,
                'ID': note_id,
                'TASK_ID': last_task_id,
                'TYPE': 'note',
                'COMMENT': str(row['CONTENT']),
                'AUTHOR': row['AUTHOR'],
                'DATE': row['DATE'],
                'IND_DATE': ind_date,
                'EXACT_DATE': exact_date_str,
                'IND_PAST': ind_past
            })

            log.debug("Processed NOTE: %s, COMMENT: %s", note_id, row['CONTENT'])

        # Non-verbose log every 100 rows
        if not verbose and idx > 0 and idx % 100 == 0:
            log.info("Processed %d rows from %s", idx, filepath)

    # Display the total number of rows processed in the file
    log.info("Finished processing %s. Total rows processed: %d", filepath, len(df))

    return task_id_counter, note_id_counter


def save_output(final_df, output_filepath):
    while True:
        try:
            # Attempt to save the final DataFrame to an Excel file
            final_df.to_excel(output_filepath, index=False)
            log.info("Migration complete. Data saved to %s", output_filepath)
            break
        except PermissionError:
            log.error("Error: Unable to save to %s. The file may be open.", output_filepath)
            retry = input("Would you like to retry saving? (Y/N): ").strip().upper()
            if retry == "N":
                log.info("Exiting without saving.")
                break


def main(verbose=False):
    df_list = []
    task_id_counter = 1  # Start the task ID counter from 1
    note_id_counter = 1  # Start the note ID counter from 1

    # Walk through all files in the source directory
    for root, dirs, files in os.walk(HARDCODED_SOURCE_FOLDER):
        for file in files:
            if file.endswith('.csv'):
                filepath = os.path.join(root, file)
                task_id_counter, note_id_counter = process_file(filepath, df_list, task_id_counter, note_id_counter,
                                                                verbose)

    # Convert the list of dictionaries to a DataFrame
    final_df = pd.DataFrame(df_list)

    # Define the output file path
    output_filepath = os.path.join(HARDCODED_SOURCE_FOLDER, 'Todoist_Migration.xlsx')

    # Attempt to save the output file with error handling
    save_output(final_df, output_filepath)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    # Ask the user if they want verbose processing (case-insensitive Y/N)
    verbose_input = input("Do you want verbose processing? (Y/N): ").strip().upper()
    verbose = verbose_input == "Y"

    main(verbose)
