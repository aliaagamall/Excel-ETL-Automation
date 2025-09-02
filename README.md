# Excel ETL Automation with Python
This repository contains a Python script that automates an **Extract, Transform, Load (ETL)** process for processing sales data stored in an Excel file (`sales_data.xlsx`). The script extracts raw data, transforms it by cleaning and enriching it, and loads the processed data back into the Excel file with formatting for better readability and analysis.
![](https://github.com/aliaagamall/Excel-ETL-Automation/blob/main/Demo.gif)
## Features
- **Extract**: Reads raw sales data from the "Raw Data" sheet of an Excel file.
- **Transform**:
  - Cleans data by handling missing values, converting data types, and removing duplicates.
  - Adds calculated columns: `Revenue` (Quantity × Price) and `Profit` (Revenue - Cost × Quantity).
  - Combines new data with existing data in the "Processed Data" sheet, ensuring no duplicates.
  - Sorts data by date in descending order.
- **Load**: Writes the processed data to the "Processed Data" sheet in the Excel file.
- **Formatting**:
  - Converts the output data into an Excel table for better usability.
  - Applies currency formatting to `Price`, `Revenue`, and `Profit` columns.
  - Adds conditional formatting to the `Profit` column with a red-to-green gradient based on profit values.
  - Auto-adjusts column widths for readability.
## File Structure
- `sales_data.xlsx`: The input Excel file containing the raw data and where processed data will be saved.
- `ETL Script.ipynb`:  Performs the ETL process and formatting.
## Usage
1. Ensure `sales_data.xlsx` is in the same directory as the script and contains a "Raw Data" sheet with the required columns.
2. The script will process the data and save the results in the "Processed Data" sheet of `sales_data.xlsx` with applied formatting.
## Output
- The processed data is saved in the "Processed Data" sheet of `sales_data.xlsx`.
- The sheet is formatted as an Excel table with:
  - Currency formatting for `Price`, `Revenue`, and `Profit`.
  - A red-to-green gradient for the `Profit` column based on values.
  - Auto-adjusted column widths for better readability.
## Notes
- Ensure the input Excel file (`sales_data.xlsx`) exists and has a "Raw Data" sheet with the expected columns.
- If the "Processed Data" sheet already exists, the script appends new data, removes duplicates, and sorts by date.
- The script uses `openpyxl` for Excel formatting, so ensure it is installed.
