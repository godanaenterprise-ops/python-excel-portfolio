
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.chart import BarChart, LineChart, Reference

print("BUSINESS INTELLIGENCE SYSTEM")
print("=" * 40)

# Sample business transactions
data = [
    ["2026-01-10", "Sales", "Income", 50000],
    ["2026-01-15", "Transport", "Expense", 8000],
    ["2026-02-10", "Sales", "Income", 65000],
    ["2026-02-15", "Food", "Expense", 12000],
    ["2026-03-10", "Sales", "Income", 75000],
    ["2026-03-15", "Rent", "Expense", 15000],
    ["2026-04-10", "Sales", "Income", 85000],
    ["2026-04-15", "Transport", "Expense", 10000],
    ["2026-05-10", "Sales", "Income", 95000],
    ["2026-05-15", "Food", "Expense", 18000],
    ["2026-06-10", "Sales", "Income", 110000],
    ["2026-06-15", "Rent", "Expense", 20000],
]

df = pd.DataFrame(
    data,
    columns=["Date", "Category", "Type", "Amount"]
)

df["Date"] = pd.to_datetime(df["Date"])

# Monthly financial analysis
df["Month"] = df["Date"].dt.to_period("M").astype(str)

monthly = df.pivot_table(
    index="Month",
    columns="Type",
    values="Amount",
    aggfunc="sum",
    fill_value=0
)

# Ensure both columns exist
for col in ["Income", "Expense"]:
    if col not in monthly.columns:
        monthly[col] = 0

monthly["Profit"] = (
    monthly["Income"] - monthly["Expense"]
)

monthly["Profit Margin (%)"] = (
    monthly["Profit"] / monthly["Income"] * 100
).fillna(0)

# Category analysis
category_report = df.pivot_table(
    index="Category",
    columns="Type",
    values="Amount",
    aggfunc="sum",
    fill_value=0
)

# Simple next-month revenue estimate
monthly = monthly.sort_index()

forecast = monthly["Income"].tail(3).mean()

forecast_report = pd.DataFrame({
    "Metric": ["Next month estimated revenue"],
    "Amount (KSh)": [forecast]
})

# Overall performance
total_income = df.loc[
    df["Type"] == "Income", "Amount"
].sum()

total_expenses = df.loc[
    df["Type"] == "Expense", "Amount"
].sum()

print("\nBUSINESS PERFORMANCE")
print(f"Total Income: KSh {total_income:,.2f}")
print(f"Total Expenses: KSh {total_expenses:,.2f}")
print(f"Total Profit: KSh {total_income-total_expenses:,.2f}")
print(f"Estimated Next Month Revenue: KSh {forecast:,.2f}")

# Export Excel workbook
filename = "Business_Intelligence_Report.xlsx"

with pd.ExcelWriter(filename, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Transactions", index=False)
    monthly.to_excel(writer, sheet_name="Monthly Analysis")
    category_report.to_excel(writer, sheet_name="Categories")
    forecast_report.to_excel(writer, sheet_name="Forecast", index=False)

# Format workbook and add charts
workbook = load_workbook(filename)

for ws in workbook.worksheets:
    ws.freeze_panes = "A2"

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="17365D"
        )

    for column in ws.columns:
        width = max(
            len(str(cell.value or "")) for cell in column
        )
        ws.column_dimensions[
            column[0].column_letter
        ].width = min(width + 4, 30)

# Monthly performance chart
ws = workbook["Monthly Analysis"]

chart = BarChart()
chart.title = "Monthly Income and Expenses"
chart.y_axis.title = "Amount (KSh)"
chart.x_axis.title = "Month"

chart_data = Reference(
    ws, min_col=2, max_col=3,
    min_row=1, max_row=ws.max_row
)

categories = Reference(
    ws, min_col=1, min_row=2,
    max_row=ws.max_row
)

chart.add_data(chart_data, titles_from_data=True)
chart.set_categories(categories)
chart.height = 9
chart.width = 18

ws.add_chart(chart, "G2")

# Profit trend chart
profit_chart = LineChart()
profit_chart.title = "Monthly Profit Trend"
profit_chart.y_axis.title = "Profit (KSh)"

profit_data = Reference(
    ws, min_col=4, min_row=1,
    max_row=ws.max_row
)

profit_chart.add_data(
    profit_data, titles_from_data=True
)

profit_chart.set_categories(categories)
profit_chart.height = 9
profit_chart.width = 18

ws.add_chart(profit_chart, "G20")

workbook.save(filename)

print("\n" + "=" * 40)
print("BUSINESS REPORT CREATED SUCCESSFULLY!")
print("File:", filename)
