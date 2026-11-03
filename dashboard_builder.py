
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter

filename = "Financial_Report.xlsx"

# Read existing financial records
df = pd.read_excel(filename, sheet_name="Transactions")

# Calculate financial performance
income = df.loc[df["Type"] == "Income", "Amount"].sum()
expenses = df.loc[df["Type"] == "Expense", "Amount"].sum()
profit = income - expenses

# Group expenses by category
expense_data = df[df["Type"] == "Expense"]

category_summary = (
    expense_data.groupby("Category")["Amount"]
    .sum()
    .reset_index()
)

# Open existing Excel workbook
workbook = load_workbook(filename)

# Replace dashboard if it already exists
if "Dashboard" in workbook.sheetnames:
    del workbook["Dashboard"]

sheet = workbook.create_sheet("Dashboard", 0)

# Dashboard title
sheet.merge_cells("A1:E1")
sheet["A1"] = "FINANCIAL PERFORMANCE DASHBOARD"
sheet["A1"].font = Font(
    bold=True, size=18, color="FFFFFF"
)
sheet["A1"].fill = PatternFill(
    "solid", fgColor="17365D"
)
sheet["A1"].alignment = Alignment(
    horizontal="center"
)

# Financial KPIs
sheet["A3"] = "Financial Metric"
sheet["B3"] = "Amount (KSh)"

metrics = [
    ("Total Income", income),
    ("Total Expenses", expenses),
    ("Net Profit", profit)
]

for row, (name, amount) in enumerate(metrics, start=4):
    sheet.cell(row, 1, name)
    sheet.cell(row, 2, float(amount))
    sheet.cell(row, 2).number_format = '"KSh "#,##0.00'

# Format KPI headers
for cell in sheet[3]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(
        "solid", fgColor="4472C4"
    )

# Expense category table
sheet["D3"] = "Expense Category"
sheet["E3"] = "Amount (KSh)"

for cell in sheet[3][3:5]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(
        "solid", fgColor="4472C4"
    )

for row, item in enumerate(
    category_summary.itertuples(index=False),
    start=4
):
    sheet.cell(row, 4, item[0])
    sheet.cell(row, 5, float(item[1]))
    sheet.cell(row, 5).number_format = '"KSh "#,##0.00'

# Income versus expenses chart
bar_chart = BarChart()
bar_chart.title = "Financial Performance"
bar_chart.y_axis.title = "Amount (KSh)"
bar_chart.x_axis.title = "Metric"

bar_data = Reference(
    sheet, min_col=2, min_row=3, max_row=6
)
bar_categories = Reference(
    sheet, min_col=1, min_row=4, max_row=6
)

bar_chart.add_data(bar_data, titles_from_data=True)
bar_chart.set_categories(bar_categories)
bar_chart.height = 8
bar_chart.width = 15

sheet.add_chart(bar_chart, "A9")

# Expense category pie chart
if len(category_summary) > 0:
    pie_chart = PieChart()
    pie_chart.title = "Expense Breakdown"

    pie_data = Reference(
        sheet,
        min_col=5,
        min_row=3,
        max_row=3 + len(category_summary)
    )

    pie_labels = Reference(
        sheet,
        min_col=4,
        min_row=4,
        max_row=3 + len(category_summary)
    )

    pie_chart.add_data(
        pie_data, titles_from_data=True
    )
    pie_chart.set_categories(pie_labels)
    pie_chart.height = 8
    pie_chart.width = 12

    sheet.add_chart(pie_chart, "J9")

# Adjust column widths
for column in range(1, 6):
    sheet.column_dimensions[
        get_column_letter(column)
    ].width = 23

# Freeze headers
sheet.freeze_panes = "A3"

# Save workbook
workbook.save(filename)

print("================================")
print("FINANCIAL DASHBOARD CREATED!")
print("================================")
print(f"Income: KSh {income:,.2f}")
print(f"Expenses: KSh {expenses:,.2f}")
print(f"Profit: KSh {profit:,.2f}")
print(f"Saved to: {filename}")
