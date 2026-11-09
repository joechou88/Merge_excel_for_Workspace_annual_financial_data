import os
import re
import logging
import warnings
import pandas as pd
from datetime import datetime
import config

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_filename = f"3_country_integrate_{timestamp}.txt"

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
target_suffix = f"_{config.START_YEAR}-{config.END_YEAR}.xlsx"
country_code_file = './country-code.xlsx'

logging.info(f"Starting country integration... Log saved to {log_filename}")

# Load country codes mapping
code_mapping = {}
if os.path.exists(country_code_file):
    logging.info(f"Loading country codes from {country_code_file}")
    code_df = pd.read_excel(country_code_file)
    for index, row in code_df.iterrows():
        country_name = str(row.get('Country_name', '')).strip()
        if country_name:
            normalized_key = country_name.lower().replace("-", " ")
            code_mapping[normalized_key] = {
                'Country_code': row.get('Country_code', ''),
                'Country_code2': row.get('Country_code2', '')
            }
else:
    logging.warning(f"{country_code_file} not found. Country codes will be empty.")

files_to_process = []

# Scan for integrated year files (e.g., Argentina_2012-2025.xlsx)
for file_name in os.listdir(base_directory):
    if not file_name.endswith(target_suffix) or file_name.startswith('~'):
        continue
        
    base_prefix = file_name.replace(target_suffix, "")
    
    # Extract country name and optional company number (e.g., Germany1 -> Germany, 1)
    match = re.match(r"^([A-Za-z\-]+)(\d*)$", base_prefix)
    if match:
        country = match.group(1)
        company_num = int(match.group(2)) if match.group(2) else 0
        
        files_to_process.append({
            "file_name": file_name,
            "filepath": os.path.join(base_directory, file_name),
            "country": country,
            "company_num": company_num
        })

if not files_to_process:
    logging.info("No target files found to integrate.")
else:
    # Sort alphabetically by country, then sequentially by company number 
    # This ensures Germany1 comes before Germany2, preserving the requested order
    files_to_process.sort(key=lambda x: (x["country"], x["company_num"]))
    
    dataframes = []
    unique_countries = set()
    
    for file_info in files_to_process:
        logging.info(f" Reading: {file_info['file_name']}")
        df = pd.read_excel(file_info["filepath"])
        
        country_name = file_info["country"]
        search_key = country_name.lower().replace("-", " ")
        country_code = code_mapping.get(search_key, {}).get('Country_code', '')
        country_code2 = code_mapping.get(search_key, {}).get('Country_code2', '')
        
        display_country_name = country_name.replace("-", " ")
        unique_countries.add(display_country_name)

        # Add required country identifier columns
        df['country'] = display_country_name
        df['country_code'] = country_code
        df['country_code2'] = country_code2

        df['_company_num'] = file_info["company_num"]
        
        dataframes.append(df)

    logging.info("Merging all datasets...")
    # pd.concat automatically aligns columns and skips the headers of subsequent dataframes
    merged_df = pd.concat(dataframes, ignore_index=True)

    # Sort by country, year, and company number
    if 'year' in merged_df.columns:
        merged_df.sort_values(by=['country', 'year', '_company_num'], inplace=True)
    else:
        merged_df.sort_values(by=['country', '_company_num'], inplace=True)
        
    # Remove temporary sort column
    merged_df.drop(columns=['_company_num'], inplace=True)
    
    desired_front = ['Type', 'year', 'country', 'country_code', 'country_code2']
    front_columns = [col for col in desired_front if col in merged_df.columns]
    other_columns = [col for col in merged_df.columns if col not in front_columns]
    merged_df = merged_df[front_columns + other_columns]

    num_countries = len(unique_countries)
    output_filename = f"All{num_countries}Countries_{config.START_YEAR}-{config.END_YEAR}.csv"
    output_filepath = os.path.join(base_directory, output_filename)
    
    logging.info(f"Saving merged data to {output_filename}")
    try:
        merged_df.to_csv(output_filepath, index=False, encoding='utf-8-sig')
        logging.info("--- Country integration completed ---")
    except Exception as e:
        logging.error(f"Failed to save {output_filename}: {e}")
