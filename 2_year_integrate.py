import os
import logging
import warnings
import openpyxl
import pandas as pd
from datetime import datetime
import config

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_filename = f"2_year_integrate_{timestamp}.txt"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(log_filename, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

base_directory = './raw_data'
file_groups = {}

logging.info(f"Starting year integration... Log saved to {log_filename}")

target_input = input("Enter 'Country' to start from (alphabetically, e.g., Mexico) or leave blank for all: ").strip().lower()

# Group files by base prefix (country + optional company identifier)
for root_directory, directories, files in os.walk(base_directory):
    for file_name in files:
        if not file_name.endswith('.xlsm') or file_name.startswith('~'):
            continue
            
        match = config.FILENAME_PATTERN.match(file_name)
        if not match:
            continue

        # Skip files with suffixes (e.g., _A) since they were processed in step 1
        if match.group('suffix'):
            continue

        country = match.group('country') or match.group('category')
        company = match.group('company') or ""
        base_prefix = f"{country}{company}"

        if base_prefix not in file_groups:
            file_groups[base_prefix] = []
            
        filepath = os.path.join(root_directory, file_name)
        file_groups[base_prefix].append(filepath)

sorted_prefixes = sorted(file_groups.keys())

for base_prefix in sorted_prefixes:
    # Skip if the alphabetical order precedes the target
    if target_input and base_prefix.lower() < target_input:
        continue

    filepaths = file_groups[base_prefix]
    logging.info(f"Processing group: {base_prefix} ({len(filepaths)} files)")
    
    dataframes = []
    base_xlsm_path = filepaths[0]  # Used as template to retain VBA macros
    
    for filepath in filepaths:
        file_name = os.path.basename(filepath)
        match = config.FILENAME_PATTERN.match(file_name)
        
        start_year = int(match.group('start_year'))
        end_year = int(match.group('end_year')) if match.group('end_year') else start_year
        is_multi_year = bool(match.group('end_year'))

        # Skip file if its entire range is outside the target period
        if end_year < config.START_YEAR or start_year > config.END_YEAR:
            logging.info(f"  Skipping {file_name} (out of target year range)")
            continue

        logging.info(f"  Reading: {file_name}")

        if is_multi_year:
            # Read REQUEST_TABLE to map years to their corresponding sheets
            workbook_source = openpyxl.load_workbook(filepath, data_only=True)
            if "REQUEST_TABLE" in workbook_source.sheetnames:
                sheet_request = workbook_source["REQUEST_TABLE"]
                
                for row in sheet_request.iter_rows(min_row=7, min_col=7, max_col=11, values_only=True):
                    year_value = row[0]  # Column G
                    cell_ref = row[4]    # Column K
                    
                    if year_value and cell_ref:
                        try:
                            year_value = int(year_value)
                        except ValueError:
                            continue
                            
                        if config.START_YEAR <= year_value <= config.END_YEAR:
                            sheet_name = str(cell_ref).split('!')[0].replace("'", "")
                            if sheet_name in workbook_source.sheetnames:
                                df = pd.read_excel(filepath, sheet_name=sheet_name)
                                df['year'] = year_value
                                dataframes.append(df)
            workbook_source.close()
        else:
            # Single-year file, data expected in '工作表1'
            if config.START_YEAR <= start_year <= config.END_YEAR:
                try:
                    df = pd.read_excel(filepath, sheet_name="工作表1")
                    df['year'] = start_year
                    dataframes.append(df)
                except ValueError:
                    logging.warning(f"  '工作表1' not found in {file_name}")
                    
    if not dataframes:
        logging.info(f"  No valid data found for {base_prefix}")
        continue

    # Merge datasets vertically, sort by year then original company order
    merged_df = pd.concat(dataframes, ignore_index=True)
    primary_col = merged_df.columns[0]
    original_type_order = merged_df[primary_col].dropna().drop_duplicates().tolist()
    merged_df[primary_col] = pd.Categorical(merged_df[primary_col], categories=original_type_order, ordered=True)
    merged_df.sort_values(by=['year', primary_col], inplace=True)
    
    # Move 'year' column to be immediately after 'Type' column
    other_columns = [col for col in merged_df.columns if col not in [primary_col, 'year']]
    columns = [primary_col, 'year'] + other_columns
    merged_df = merged_df[columns]

    # Filter out empty columns automatically named 'Unnamed: xx' by Pandas
    merged_df = merged_df.loc[:, ~merged_df.columns.str.contains('^Unnamed')]
    
    # Convert datetime objects to string format 'YYYY/MM/DD' to handle mixed data types (e.g., dates with '.')
    for column in merged_df.columns:
        if pd.api.types.is_datetime64_any_dtype(merged_df[column]):
            merged_df[column] = merged_df[column].dt.strftime('%Y/%m/%d')
        elif merged_df[column].dtype == 'object':
            merged_df[column] = merged_df[column].apply(
                lambda val: val.strftime('%Y/%m/%d') if isinstance(val, (pd.Timestamp, datetime)) and pd.notna(val) else val
            )

    output_filename = f"{base_prefix}_{config.START_YEAR}-{config.END_YEAR}.xlsx"
    output_filepath = os.path.join(base_directory, output_filename)
    
    logging.info(f"  Saving merged file to {output_filename}")
    
    try:
        merged_df.to_excel(output_filepath, index=False, sheet_name="工作表1")
    except Exception as e:
        logging.error(f"  Failed to save {output_filename}: {e}")

logging.info("--- Year integration completed ---")
