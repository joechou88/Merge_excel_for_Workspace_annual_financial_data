import os
import logging
import warnings
import openpyxl
import config
from datetime import datetime

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_filename = f"1_variable_integrate_{timestamp}.txt"

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

logging.info(f"Starting variable integration... Log saved to {log_filename}")

# Prompt user for specific target
target_input = input("Enter 'Country, Year' to run specifically (e.g., Australia, 2012) or leave blank for all: ").strip()
target_country, target_year = None, None
if target_input and ',' in target_input:
    target_country, target_year = [p.strip() for p in target_input.split(',')]

# Group files by their base name, skipping files without variable group suffixes
for root_directory, directories, files in os.walk(base_directory):
    for file_name in files:
        if not file_name.endswith('.xlsm') or file_name.startswith('~'):
            continue
            
        match = config.FILENAME_PATTERN.match(file_name)
        if not match:
            continue

        # Filter by specific country and year if provided by user
        if target_country and target_year:
            file_country = match.group('country') or match.group('category')
            file_year = match.group('start_year')
            if file_country != target_country or file_year != target_year:
                continue
            
        suffix = match.group('suffix')
        if not suffix:
            # Skip files without variable group suffixes (e.g., Ireland_2017-2018.xlsm)
            continue
            
        # Create the final merged file name (e.g., Germany1_2012_A.xlsm -> Germany1_2012.xlsm)
        base_file_name = file_name.replace(f"_{suffix}.xlsm", ".xlsm")
        group_key = os.path.join(root_directory, base_file_name)
        
        if group_key not in file_groups:
            file_groups[group_key] = {
                'files': [],
                'is_multi_year': bool(match.group('end_year'))
            }
            
        filepath = os.path.join(root_directory, file_name)
        file_groups[group_key]['files'].append(filepath)

for group_key, group_data in file_groups.items():
    filepaths = sorted(group_data['files'])
    is_multi_year = group_data['is_multi_year']
    
    if not filepaths:
        continue
        
    logging.info(f"Merging into: {group_key} ({len(filepaths)} files)")
    
    base_filepath = filepaths[0]
    try:
        logging.info(f"  Using base file: {os.path.basename(base_filepath)}")
        # Load the first file as base to preserve VBA macros and REQUEST_TABLE
        workbook_base = openpyxl.load_workbook(base_filepath, keep_vba=True)
        
        target_sheets = []
        
        if is_multi_year:
            # For multi-year files, read REQUEST_TABLE to find target sheets mapped by year (G7=Year, K7=Target cell)
            sheet_request = workbook_base["REQUEST_TABLE"]
            for row in sheet_request.iter_rows(min_row=7, min_col=7, max_col=11, values_only=True):
                year_value = row[0]  # Column G
                cell_ref = row[4]    # Column K
                
                if year_value and cell_ref and '!' in str(cell_ref):
                    # Extract sheet name (supports formats like "工作表1'!A1" or "工作表1!A1")
                    sheet_name = str(cell_ref).split('!')[0].replace("'", "")
                    if sheet_name not in target_sheets and sheet_name in workbook_base.sheetnames:
                        target_sheets.append(sheet_name)
        else:
            # For single-year files, unify and append data to "工作表1"
            target_sheets = ["工作表1"]

        # Read other variable files in the group and append data horizontally
        for filepath in filepaths[1:]:
            logging.info(f"  Appending from: {os.path.basename(filepath)}")
            workbook_source = openpyxl.load_workbook(filepath, data_only=True)
            
            for sheet_name in target_sheets:
                if sheet_name not in workbook_source.sheetnames:
                    continue
                    
                sheet_source = workbook_source[sheet_name]
                sheet_base = workbook_base[sheet_name]
                
                # Build primary key (Column A) to row index mapping from base sheet
                base_key_map = {
                    sheet_base.cell(row=r, column=1).value: r 
                    for r in range(1, sheet_base.max_row + 1)
                    if sheet_base.cell(row=r, column=1).value is not None
                }
                
                base_max_col = sheet_base.max_column

                # Filter out empty columns by checking the header row (row 1)
                header_row = next(sheet_source.iter_rows(min_row=1, max_row=1, values_only=True), [])
                valid_col_indices = [
                    i for i, val in enumerate(header_row) 
                    if i > 0 and val is not None and str(val).strip() != ""
                ]
                
                # Iterate all rows (including header) to append columns horizontally
                for row in sheet_source.iter_rows(values_only=True):
                    primary_key = row[0]
                    if primary_key in base_key_map:
                        target_row = base_key_map[primary_key]
                        # Append from the 2nd column onwards to the right
                        for c_idx, val in enumerate(row[1:], start=1):
                            sheet_base.cell(row=target_row, column=base_max_col + c_idx, value=val)
                        
            workbook_source.close()

        # Remove the previous merged file if it exists to ensure clean overwrite
        if os.path.exists(group_key):
            os.remove(group_key)

        # Save as the merged filename (removing suffixes like _A)
        workbook_base.save(group_key)
        workbook_base.close()
        
    except Exception as e:
        logging.error(f"Error merging {group_key}: {e}")

logging.info("--- Variable integration completed ---")
