## Integrate variable groups
The script consolidates `.xlsm` files that are split by variable groups (e.g., `Germany1_2012_A.xlsm`, `_B`, `_C`) into a single merged file per country/category and year. The `_A` file acts as the base to preserve VBA macros and the `REQUEST_TABLE`. Data from subsequent files is horizontally appended using Column A (`Type`) as the primary key, cleanly overwriting any previously generated merged files.

- **Multi-Year Sheet Mapping:** For multi-year files, the script dynamically reads the base `REQUEST_TABLE` to route data to the correct target sheet for each year. Single-year files default to merging data directly onto `工作表1`.
- **Data Alignment & Filtering:** Empty columns in source files are automatically filtered out via header-row validation to prevent data misalignment during the horizontal merge. 
- **Targeted Execution:** You can run the integration across the entire directory automatically, or provide a specific target (e.g., `Australia, 2012`) via the terminal prompt to quickly update a single dataset.

## Integrate multi-year data within same country
The script consolidates `.xlsm` files that are split across different timeframes (e.g., `Belgium_2010-2014.xlsm`, `Belgium_2015.xlsm`) into a single `.xlsx` dataset per country/category covering a specified target period. The script merges datasets vertically, sorts by year then original company order, and shifts a dedicated `year` column to the first position for a cleaner format.

- **Dynamic Year Extraction:** For multi-year source files, the script dynamically reads the `REQUEST_TABLE` to map and extract target years from their corresponding sheets. Single-year files extract data directly from `工作表1`. Any years falling outside the configured `START_YEAR` and `END_YEAR` range are automatically skipped.
- **Data Cleaning & Formatting:** The script automatically filters out ghost columns (e.g., `Unnamed: xx`) caused by trailing empty formats in the original Excel files. It also standardizes mixed datetime types and missing values by forcing valid dates into a strict `YYYY/MM/DD` string format, preventing `0:00:00` artifacts and `NaT` conversion errors.
- **Targeted Execution:** You can run the integration across the entire directory automatically, or provide a specific base prefix (e.g., `Argentina` or `Germany1`) via the terminal prompt to quickly update a single dataset.

## Integrate all country datasets
The script consolidates individual country-level `.xlsx` files (e.g., `Argentina_2012-2025.xlsx`, `Germany1_2012-2025.xlsx`) into a single comprehensive `.csv` master file encompassing all countries for the specified timeframe. It merges datasets vertically, appends standardized country codes, applies strict hierarchical sorting, and dynamically shifts primary identifiers to the front.

- **Country Code Mapping:** The script reads `country-code.xlsx` to automatically append `country_code` and `country_code2` columns. It normalizes names dynamically (e.g., replacing hyphens with spaces and ignoring case), ensuring that filenames like `New-Zealand` perfectly match the lookup table and output as a clean `New Zealand`.
- **Advanced Sorting:** Multi-part country datasets (e.g., `Germany1` through `Germany8`) are intelligently grouped and sequenced rather than blindly appended. The final master file is sorted hierarchically by `country`, `year`, and the original company sequence to maintain logical continuity.
- **Format Optimization & Export:** Essential identifier columns (`Type`, `year`, `country`, `country_code`, `country_code2`) are automatically repositioned to the very front of the dataset for optimal readability. The result is exported with a dynamic filename counting the total unique countries (e.g., `All41Countries_2012-2025.csv`).

## Standardize Column Names
The script processes the consolidated multi-country `All(num)Countries_(START_YEAR)-(END_YEAR).csv` file to standardize column headers and remove redundant data. It reads the dataset, identifies and drops duplicated columns generated during previous merges, and applies regular expressions to clean up messy financial column codes before exporting the refined dataset as a new `All(num)Countries_(START_YEAR)-(END_YEAR)_renamed.csv` file.

| Original Column Name | Converted Name | Description |
|----------------------|----------------|-------------|
| Type                 | DSCD           | Fixed renaming |
| X(WC01254)           | WC01254        | Remove outer `X()` |
| X(WC06705)~U         | XWC06705U      | Remove parentheses & `~`, keep suffix |
| X(WC02051)~U$.1      | XWC02051U      | Remove the entire column, as the field is repetitive. |
| X(WC18545)~U$        | XWC18545U      | Remove `~` and `$` |
| X(WC04601)~US        | XWC04601U      | Standardize suffix to a single `U` |

- **Regex-Based Column Renaming:** Automatically identifies and simplifies complex column headers using regex patterns based on the rules in the table above, without hardcoding specific item codes. 
- **Duplicate Column Removal:** It automatically detects and drops duplicate columns generated by pandas (e.g., those ending with `.1`, `.2`), ensuring a clean and non-redundant dataset.
- **Detailed Excel-Style Logging:** Provides comprehensive runtime logging that converts numeric column indices into intuitive Excel-style column letters (e.g., A, B, AA). This allows users to easily track exactly which columns were dropped or renamed, mapping the programmatic changes directly back to the original spreadsheet layout.
