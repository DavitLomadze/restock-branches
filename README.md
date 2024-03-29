# Request Form preparation script
This script is designed to automate the preparation of Excel request forms for branch managers in a retail setting, allowing them to order products for their branches efficiently. It involves reading and processing product evaluation, sales, and inventory data from CSV and Excel files, then generating detailed Excel reports with product recommendations for each branch. The script employs the pandas, numpy, and openpyxl libraries for data manipulation and Excel file operations, and uses logging for tracking its execution process.

## Overview
- Purpose: Automate generation of Excel request forms for branch managers.
- Input Data:
- Product evaluations (product_evaluation.csv)
- Sales data (sales_cleaned.csv)
- Inventory data (inventory_clean.csv)
- Closing inventory with margins (closing_inventory_margins.xlsx)
- Product descriptions (product_description.xlsx)
- Codes to be removed (remove_codes.xlsx)
- Output: Excel files for each branch with product order recommendations.

## Key Functions:
### remove_codes(code_dir: str) -> pd.DataFrame
Removes specified product codes from consideration.

Parameters:
- code_dir: Directory containing the Excel file with codes to remove.
- Returns: DataFrame with codes to remove.

### prep_dataframes(...)
Prepares and cleans the data from the provided CSV and Excel files for further analysis.

Returns: Multiple DataFrames containing cleaned and structured data for analysis.

### request_form(...)
Generates a detailed DataFrame for each warehouse, containing product recommendations based on inventory, sales, and product evaluation data.

Parameters: Includes warehouse information, data frames prepared by prep_dataframes, and other necessary data.
Returns: A DataFrame ready to be converted into an Excel file for the branch.

### calculate_last_row(dataframe)
Calculates the last row number for an Excel table based on the DataFrame size.

Parameters:
dataframe: DataFrame to calculate the last row for.
Returns: The last row number as an integer.

### initiate_excel_file()
Initializes a new Excel workbook and sheet.

Returns: The active worksheet and workbook objects.

### format_excel_file(...)
Applies formatting to the generated Excel file, including table formatting, header styling, and data validation.

Parameters: Worksheet, last row number, and warehouse information.
### populate_excel_file(...)
Fills the Excel file with data from the DataFrame generated by request_form.

Parameters: Worksheet, last row, the DataFrame with details, inventory DataFrame, and warehouse information.

### save_excel_file(...)
Saves the Excel workbook to a specified location.

Parameters:
wb: Workbook object to save.
warehouse: Warehouse information to name the file appropriately.
Execution
The script executes by calling the main function, which sequentially prepares data frames, generates request forms for each warehouse, formats the Excel files, and saves them. It employs exception handling to manage errors during execution and uses logging to record the process and any issues encountered.

## Logging
Logging is set up at the beginning of the script to track its execution and troubleshoot any problems. It logs both standard operation messages and exceptions.
