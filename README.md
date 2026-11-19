# Formatted Excel Sales Report

## Project Overview
This project automates the generation of a professional Excel sales report using Python.

The report analyses wholesale sales transactions and generates an Excel workbook containing business performance indicators, raw sales data, and an embedded revenue chart.

## Features
- Reads sales data from a CSV file.
- Calculates total revenue, quantity sold, transaction count, and average revenue.
- Summarizes revenue by product.
- Creates Summary and Data worksheets.
- Applies professional Excel formatting.
- Embeds a bar chart showing revenue by product.

## Technologies
- Python
- Pandas
- OpenPyXL

## How to Run

1. Install the required libraries:

   `pip install pandas openpyxl`

2. Place `sales_data.csv` in the project folder.

3. Run:

   `python generate_report.py`

4. Open the generated `sales_report.xlsx` workbook in Microsoft Excel.

## Project Files
- `generate_report.py` – Report generation script.
- `sales_data.csv` – Raw sales dataset.
- `sales_report.xlsx` – Generated Excel report.

## Author
Python Data Analysis Portfolio Project.