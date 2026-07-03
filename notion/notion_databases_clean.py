import logging
import os
import sys
import pandas as pd
import re
import warnings
import ast

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
log = logging.getLogger(__name__)

# requirements: custom functions
from utils import read_params_from_txt_file, get_column_widths, apply_column_widths, DEFAULT_PARAMS_FILE

# Suppress openpyxl warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet._reader")

# Ensure stdout can emit unicode regardless of console code page
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except AttributeError:
    pass

def read_database_list(excel_path):
    df = pd.read_excel(excel_path)
    return df[df["ind_clean"] == 1]

def load_metadata(metadata_path):
    log.info("📂 Loading metadata from: %s", metadata_path)
    return pd.read_excel(metadata_path)

def clean_data(row, col_value, row_num=None):
    database = row['database']
    col = row['column']
    clean = row['clean']
    extract = row['extract']
    between = row['between']
    json_check = row['json']
    dictionary1 = row['dictionary1']
    dictionary2 = row['dictionary2']

    # Define a function to extract content from JSON nested dictionaries
    def extract_content(json_str, dictionary1, dictionary2, row_num):
        try:
            json_data = ast.literal_eval(json_str)
        except (ValueError, SyntaxError) as e:
            log.error("❌ Failed to parse JSON string at row %s: %s. Error: %s", row_num, json_str, e)
            return None

        if len(json_data) > 0 and dictionary1 in json_data[0] and dictionary2 in json_data[0][dictionary1]:
            return json_data[0][dictionary1][dictionary2]
        else:
            if verbose:
                log.debug("⚠️ Skipping row number %s due to missing data: %s", row_num, json_str)
            return None

    if json_check == 1:  # Check if the 'json' value is 1.
        extracted_value = extract_content(col_value, dictionary1, dictionary2, row_num)
        if extracted_value is not None:
            return extracted_value
        else:
            return col_value  # Return the original value if extraction failed
    elif clean == 1:
        if between == "it's a number":
            # Extract the value immediately after the specified key path and treat it as a number
            key_path = re.escape(extract)
            pattern = fr"{key_path}\s*([^,}}]+)"
            col_value_str = str(col_value)
            match = re.search(pattern, col_value_str)
            if verbose:
                log.debug("Pattern: %s, Column Value: %s, Match: %s", pattern, col_value_str, match.group(1) if match else 'None')
            return float(match.group(1)) if match else col_value
        else:
            # Original logic for other cases
            pattern = fr"{re.escape(extract)}\s*{re.escape(between)}(.*?){re.escape(between)}"
            col_value_str = str(col_value)
            match = re.search(pattern, col_value_str)
            if verbose:
                log.debug("Pattern: %s, Column Value: %s, Match: %s", pattern, col_value_str, match.group(1) if match else 'None')
            return match.group(1) if match else col_value
    else:
        return col_value

def process_databases(databases_to_process, metadata):
    log.info("📊 Databases to process: %d", len(databases_to_process))

    for _, database in databases_to_process.iterrows():
        input_path = database["output_path"]
        log.info("🔄 Processing database: %s", database['name'])
        log.info("📂 Input path: %s", input_path)
        df = pd.read_excel(input_path)

        # Define a list to hold columns to be dropped
        cols_to_drop = []

        # Dictionary to hold final column names and their order
        final_col_order = {}

        # List to keep track of columns to keep
        columns_to_keep = []

        for col in df.columns:
            metadata_row = metadata.loc[(metadata["database"] == database["name"]) & (metadata["column"] == col)]

            # If there's metadata for this column
            if not metadata_row.empty:
                if verbose:
                    log.debug("🧹 Cleaning column: %s", col)
                # Create a cleaned column
                df[f"{col}_clean"] = df.apply(lambda row: clean_data(metadata_row.iloc[0], row[col], row.name + 1), axis=1)
                # Replace "[]" with NaN
                df[f"{col}_clean"] = df[f"{col}_clean"].replace("[]", pd.NA)

                # Reading "ind_keep", "column_name_final" and "order_final" values from metadata
                ind_keep = metadata_row["ind_keep"].values[0]
                col_name_final = metadata_row["column_name_final"].values[0]
                order_final = metadata_row["order_final"].values[0]

                # Check "ind_keep" value and make adjustments to the dataframe based on its value
                if ind_keep == 0:
                    # If ind_keep is 0, we don't need this column or its cleaned version, so we'll mark them for deletion
                    cols_to_drop.append(col)
                    cols_to_drop.append(f"{col}_clean")
                    if verbose:
                        log.debug("🗑️ Marking for drop: %s, %s_clean", col, col)
                elif ind_keep == 1:
                    # If ind_keep is 1, we'll keep the cleaned column and drop the original one
                    # Also, we'll rename the cleaned column to the final name
                    cols_to_drop.append(col)
                    columns_to_keep.append(col_name_final)
                    df.rename(columns={f"{col}_clean": col_name_final}, inplace=True)
                    if verbose:
                        log.debug("📝 Renaming column %s_clean to %s", col, col_name_final)

                    # If "order_final" value is not null and it's an integer, store it in the dictionary
                    if pd.notnull(order_final):
                        final_col_order[col_name_final] = order_final
                        if verbose:
                            log.debug("🔢 Setting order for column %s: %s", col_name_final, order_final)
                elif ind_keep == 2:
                    # If ind_keep is 2, we keep both the original and cleaned columns
                    columns_to_keep.append(col)
                    columns_to_keep.append(f"{col}_clean")
                    if verbose:
                        log.debug("✅ Keeping both original and cleaned columns for %s", col)

            else:
                log.warning("⚠️ Skipping column: %s (no metadata found)", col)

        # Drop columns that are marked to be dropped
        log.info("🗑️ Columns to drop: %s", cols_to_drop)
        df.drop(cols_to_drop, axis=1, inplace=True)

        # Ensure we only keep the columns we have processed and want to keep
        df = df[columns_to_keep]

        log.info("📊 Columns after dropping and filtering to keep: %s", df.columns.tolist())

        # If there are valid "order_final" values, rearrange the DataFrame columns
        if final_col_order:
            log.info("🔢 Ordering columns: %s", final_col_order)
            sorted_columns = [col for col in sorted(final_col_order, key=final_col_order.get)]
            df = df[sorted_columns]
            log.info("✅ Columns after reordering: %s", df.columns.tolist())

        output_path = os.path.join(dump_path, f"database-dump-{database['name']}_clean.xlsx")

        # If the database is flagged as ind_format = 1 then gather and then apply widths to the excel file, before overwriting it again
        if database["ind_format"] == 1:
            log.info("📏 Applying column widths: %s", database['name'])
            column_widths = get_column_widths(output_path)
            log.info("📝 Writing Excel file: %s", output_path)
            df.to_excel(output_path, index=False, engine='openpyxl')
            if column_widths:
                apply_column_widths(output_path, column_widths)
            else:
                log.warning("⚠️ No reusable column widths found for %s.", database['name'])
        else:
            log.info("📝 Writing Excel file: %s", output_path)
            df.to_excel(output_path, index=False, engine='openpyxl')

        log.info("✅ Processed database '%s' saved to %s with %d rows", database['name'], output_path, len(df))

def main():
    params_file_path = DEFAULT_PARAMS_FILE
    log.info("📂 Loading parameters...")
    params = read_params_from_txt_file(params_file_path)
    log.info("✅ Parameters loaded")

    excel_path = params['excel_path']
    dump_path = params['dump_path']
    verbose = params['verbose'] == "True"  # Ensure this is checked correctly

    metadata_path = params['metadata_path']
    metadata = load_metadata(metadata_path)
    log.info("✅ Metadata loaded")

    log.info("📋 Reading database list...")
    databases_to_process = read_database_list(excel_path)
    log.info("📊 Found %d databases to process", len(databases_to_process))

    log.info("🚀 Starting database cleaning...")
    process_databases(databases_to_process, metadata)
    log.info("✅ All database cleaning completed")


if __name__ == "__main__":
    main()
