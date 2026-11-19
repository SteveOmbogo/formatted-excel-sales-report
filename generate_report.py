import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# ==========================================
# 1. LOAD THE RAW DATA
# ==========================================

df = pd.read_csv("sales_data.csv")

# Convert Date column into actual dates
df["Date"] = pd.to_datetime(df["Date"])

# Calculate revenue for each transaction
df["Revenue"] = df["Quantity"] * df["Unit_Price"]

# ==========================================
# 2. CALCULATE BUSINESS KPIs
# ==========================================

total_revenue = df["Revenue"].sum()
total_quantity = df["Quantity"].sum()
total_transactions = len(df)
average_revenue = df["Revenue"].mean()

print("SALES REPORT SUMMARY")
print("--------------------")
print(f"Total Revenue: KES {total_revenue:,.2f}")
print(f"Total Quantity Sold: {total_quantity}")
print(f"Total Transactions: {total_transactions}")
print(f"Average Transaction Revenue: KES {average_revenue:,.2f}")
# ==========================================
# 3. CREATE THE EXCEL WORKBOOK
# ==========================================

wb = Workbook()

# Create Summary worksheet
summary = wb.active
summary.title = "Summary"

# Create Data worksheet
data_sheet = wb.create_sheet("Data")

# ==========================================
# 4. FORMAT THE SUMMARY SHEET
# ==========================================

summary["A1"] = "WHOLESALE SALES PERFORMANCE REPORT"
summary["A1"].font = Font(
    bold=True,
    size=16,
    color="FFFFFF"
)
summary["A1"].fill = PatternFill(
    "solid",
    fgColor="1F4E78"
)
summary.merge_cells("A1:B1")

summary["A3"] = "BUSINESS PERFORMANCE INDICATORS"
summary["A3"].font = Font(bold=True, size=12)

# KPI table
kpis = [
    ["Metric", "Value"],
    ["Total Revenue (KES)", total_revenue],
    ["Total Quantity Sold", total_quantity],
    ["Total Transactions", total_transactions],
    ["Average Revenue (KES)", average_revenue]
]

for row in kpis:
    summary.append(row)

# Format KPI headings
for cell in summary[4]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="4472C4")

# Format KPI values
summary["B5"].number_format = '"KES" #,##0.00'
summary["B6"].number_format = '#,##0'
summary["B7"].number_format = '#,##0'
summary["B8"].number_format = '"KES" #,##0.00'

# ==========================================
# 5. CALCULATE REVENUE BY PRODUCT
# ==========================================

product_revenue = (
    df.groupby("Product")["Revenue"]
    .sum()
    .reset_index()
)

summary["D3"] = "REVENUE BY PRODUCT"
summary["D3"].font = Font(bold=True, size=12)

summary["D4"] = "Product"
summary["E4"] = "Revenue (KES)"

for cell in summary[4][3:5]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="4472C4")


for index, row in enumerate(
    product_revenue.itertuples(index=False),
    start=5
):
    summary.cell(index, 4, row.Product)
    summary.cell(index, 5, row.Revenue)
    summary.cell(index, 5).number_format = '"KES" #,##0.00'

# ==========================================
# 6. CREATE AN EMBEDDED BAR CHART
# ==========================================

chart = BarChart()

chart.type = "col"
chart.style = 10
chart.title = "Revenue by Product"
chart.y_axis.title = "Revenue (KES)"
chart.x_axis.title = "Product"
chart.height = 9
chart.width = 16

data = Reference(
    summary,
    min_col=5,
    min_row=4,
    max_row=4 + len(product_revenue)
)

categories = Reference(
    summary,
    min_col=4,
    min_row=5,
    max_row=4 + len(product_revenue)
)

chart.add_data(data, titles_from_data=True)
chart.set_categories(categories)

summary.add_chart(chart, "G3")

# ==========================================
# 7. POPULATE THE DATA WORKSHEET
# ==========================================

data_sheet.append(list(df.columns))

for row in df.itertuples(index=False, name=None):
    data_sheet.append(list(row))

# Format column headings
for cell in data_sheet[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center")

# Format date and revenue columns
for row in range(2, data_sheet.max_row + 1):
    data_sheet.cell(row, 1).number_format = "yyyy-mm-dd"
    data_sheet.cell(row, 6).number_format = '"KES" #,##0.00'
    data_sheet.cell(row, 7).number_format = '"KES" #,##0.00'

# Add borders
thin_border = Border(
    bottom=Side(style="thin", color="D9E2F3")
)

for row in data_sheet.iter_rows():
    for cell in row:
        cell.border = thin_border

# Enable filters and freeze headings
data_sheet.auto_filter.ref = data_sheet.dimensions
data_sheet.freeze_panes = "A2"

# ==========================================
# 8. ADJUST COLUMN WIDTHS
# ==========================================

for sheet in [summary, data_sheet]:
    for column in sheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)

        for cell in column:
            if cell.value is not None:
                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        sheet.column_dimensions[column_letter].width = max_length + 4

# ==========================================
# 9. SAVE THE EXCEL REPORT
# ==========================================

wb.save("sales_report.xlsx")

print("\nExcel report generated successfully!")
print("File saved as: sales_report.xlsx")