
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.chart import BarChart, Reference

# Sample business data
data = {
    "Month": ["January", "February", "March", "April", "May", "June"],
    "Income": [50000, 65000, 75000, 85000, 95000, 110000],
    "Expenses": [8000, 12000, 15000, 10000, 18000, 20000]
}

df = pd.DataFrame(data)

# Calculate profit
df["Profit"] = df["Income"] - df["Expenses"]

# Calculate totals
total_income = df["Income"].sum()
total_expenses = df["Expenses"].sum()
total_profit = df["Profit"].sum()

# Create Excel report
file_name = "Automated_Business_Report.xlsx"

with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Monthly Report", index=False)

    summary = pd.DataFrame({
        "Metric": ["Total Income", "Total Expenses", "Total Profit"],
        "Amount": [total_income, total_expenses, total_profit]
    })

    summary.to_excel(writer, sheet_name="Summary", index=False)

# Format workbook
wb = load_workbook(file_name)
ws = wb["Monthly Report"]

for cell in ws[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )

# Add income and expenses chart
chart = BarChart()
chart.title = "Monthly Income and Expenses"
chart.y_axis.title = "Amount"
chart.x_axis.title = "Month"

data_ref = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=7)
categories = Reference(ws, min_col=1, min_row=2, max_row=7)

chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(categories)

ws.add_chart(chart, "F2")

# Save final report
wb.save(file_name)

print("Automated report created successfully!")
print("Total Income:", total_income)
print("Total Expenses:", total_expenses)
print("Total Profit:", total_profit)
print("File:", file_name)
