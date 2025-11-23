# chatGPT source 2023-05-01 > https://chat.openai.com/c/de2353d2-d8c8-41c0-b9d2-8317c39c8c2c
# chatGPT source 2023-09-26 > https://chat.openai.com/c/2f7bf670-4b1b-4cd1-874d-0bb54174bddc
# chatGPT source 2023-10-08 > https://chat.openai.com/c/026c7ef9-a6d2-4304-ac37-452d1896c592 

# requirements: public
import os
import pandas as pd
from datetime import datetime, timedelta

# requirements: custom functions
from utils import read_params_from_txt_file

# Global variable to store the adjusted date
adjusted_date = None

def get_adjusted_date(timedelta_value=None):
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

def process_journal(journal_excel_path, journal_output_path, date_column, text_column, split_comma=True, file_suffix="journal"):
    """
    This function reads an Excel file, filters rows based on the date column, processes the text column, and writes the result to text file.

    Parameters:
    - journal_excel_path: Path to the Excel file.
    - journal_output_path: Path to the directory where the output file will be saved.
    - date_column: Name of the date column.
    - text_column: Name of the text column.
    - split_comma: If True, rows in the text column are split at commas. Default is True.
    - file_prefix: Prefix for the name of the output file. Default is "journal".
    """

    # Use the new function to get the adjusted date
    current_date = get_adjusted_date()

    # Load the Excel file into a pandas DataFrame
    print(f"📂 Loading Excel file: {journal_excel_path}")
    df = pd.read_excel(journal_excel_path)

    # Calculate the date of the previous Sunday
    days_since_sunday = (current_date.weekday() - 6) % 7
    previous_sunday = current_date - timedelta(days=days_since_sunday)

    # Calculate the date of the Monday before the previous Sunday
    one_week_ago = previous_sunday - timedelta(days=6)
    print(f"📅 Date range: {one_week_ago.strftime('%Y-%m-%d')} to {previous_sunday.strftime('%Y-%m-%d')}")

    # Fix potential time discrepancies by setting time to start of the day
    one_week_ago = one_week_ago.replace(hour=0, minute=0, second=0, microsecond=0)
    previous_sunday = previous_sunday.replace(hour=0, minute=0, second=0, microsecond=0)

    # Convert date column to datetime
    df[date_column] = pd.to_datetime(df[date_column])

    # Sort the DataFrame by date column, default is ascending
    df = df.sort_values(date_column)

    # Select the rows between the one week ago date and the previous sunday
    df = df[(df[date_column] >= one_week_ago) & (df[date_column] <= previous_sunday)]
    print(f"📊 Filtered {len(df)} rows for date range")

    # Select the text column
    df = df[text_column]

    if split_comma:
        # Split the text column rows on commas and expand into new DataFrame
        print("🔀 Splitting text on commas...")
        df = df.str.split(',', expand=True).stack().reset_index(drop=True)

    # Filter out rows that contain only '[]'
    df = df[df != '[]']

    # Drop duplicates while preserving the original order
    print("🔍 Removing duplicates...")
    df = df.drop_duplicates(keep='first')

    # Set max column width to None to avoid cutting off strings
    pd.set_option('display.max_colwidth', None)

    # Convert the DataFrame to a string, with rows separated by line carriages
    # Then, remove leading and trailing spaces from each line individually
    data_string = "\n".join(line.strip() for line in df.to_string(index=False, header=False).split('\n'))

    # Handle "\n" to interpret as a line carriage
    data_string = data_string.replace("\\n", "\n")

    # Handle "NaN" to interpret as "null" and skip that text altogether
    data_string = data_string.replace("NaN", "")

    # Filter out lines that are now empty due to NaN replacement
    data_string = "\n".join([line for line in data_string.split("\n") if line.strip()])

    # Create the output file name using the start and end dates
    output_file_name = f"{one_week_ago.strftime('%Y-%m-%d')} to {previous_sunday.strftime('%Y-%m-%d')}-{file_suffix}.txt"

    # Save the data string to a text file at the specified directory, using a path.join method
    output_path = os.path.join(journal_output_path, output_file_name)
    print(f"📝 Writing output file: {output_file_name}")
    with open(output_path, 'w') as file:
        file.write(data_string)

    print(f"✅ Saved: {output_path}")

# Main execution
params_file_path = r"C:\Mis Datos en Local\temporal\python\notion-params.txt"
print("📂 Loading parameters...")
params = read_params_from_txt_file(params_file_path)
print("✅ Parameters loaded")

print("\n🚀 Processing journal entries...")
process_journal(params['journal_excel_path'], params['journal_output_path'], "D_JOURNAL", "TXT_GRATITUDE", True, "gratitude")
process_journal(params['journal_excel_path'], params['journal_output_path'], "D_JOURNAL", "TXT_WORK", False, "work")
process_journal(params['journal_excel_path'], params['journal_output_path'], "D_JOURNAL", "TXT_PERSONAL", False, "personal")
process_journal(params['journal_excel_path'], params['journal_output_path'], "D_JOURNAL", "TXT_LEARN", False, "learning")
print("\n✅ All journal processing completed")