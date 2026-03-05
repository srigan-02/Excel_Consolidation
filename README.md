# Excel Sheet Consolidation using Python

## Overview
This script consolidates data from **multiple sheets in an Excel workbook into a single sheet** using **common column headers**.  
Only the columns that exist in **all sheets** will be included in the final output to ensure consistent structure.

This is useful when working with Excel files where each sheet contains similar structured data and needs to be merged for analysis.

---

## Features
- Reads all sheets from an Excel file.
- Identifies **common headers across sheets**.
- Filters each sheet to retain only common columns.
- Combines all rows into a **single consolidated dataset**.
- Exports the result to a new Excel file.

---

## How to Run

1. Place your Excel file in the project folder.  
2. Update the file path in the script if needed.  
3. Run the script:

## Example Use Case

You have an Excel file with multiple sheets:
- Sheet1
- Sheet2
- Sheet3
- Sheet4

Each sheet contains similar columns like:

Item Code | Description | Quantity | Price

The script will combine all rows from these sheets into one dataset containing only the columns common across all sheets.

---

## Output

The output file:

- consolidated_common_headers.xlsx
- Contains all rows from all sheets combined into one table.

---
## Possible Enhancements
- Adding Sheet Name as a Source Column
- Handling different header names automatically
- Creating data validation checks
- Handling missing columns dynamically
- Building a GUI for Excel consolidation

---
Author
Sriram G
