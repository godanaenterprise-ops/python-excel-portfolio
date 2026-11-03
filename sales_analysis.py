
import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill

# Sample sales data
data = {
    "Product": [
        "Laptop", "Phone", "Tablet",
        "Laptop", "Phone", "Tablet",
        "Laptop", "Phone", "Tablet"
    ],
    "Quantity": [5, 10, 7, 3, 8, 4, 6, 12, 5],
    "Unit Price": [80000, 30000, 25000, 80000, 30000, 25000,
                   80000, 30000, 25000]
}

df = pd.DataFrame(data)

# Calculate revenue
df["Revenue"] = df["Quantity"] * df["Unit Price"]

# Analyse sales by product
report = df.groupby("Product").agg({
    "Quantity": "sum",
    "Revenue": "sum"
}).reset_index()

# Calculate revenue contribution
total_revenue = report["Revenue"].sum()

report["Revenue Contribution (%)"] = (
    report["Revenue"] / total_revenue
) * 100

# Identify best-selling product by quantity
best_product = report.loc[report["Quantity"].idxmax(), "Product"]

# Export Excel report
file_name = "Sales_Analysis_Report.xlsx"

with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Sales Data", index=False)
    report.to_excel(writer, sheet_name="Product Analysis", index=False)

# Format workbook
wb = load_workbook(file_name)
ws = wb["Product Analysis"]

for cell in ws[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )

# Add revenue chart
chart = BarChart()
chart.title = "Revenue by Product"
chart.y_axis.title = "Revenue"
chart.x_axis.title = "Product"

data_ref = Reference(ws, min_col=3, min_row=1, max_row=4)
categories = Reference(ws, min_col=1, min_row=2, max_row=4)

chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(categories)

ws.add_chart(chart, "F2")

wb.save(file_name)

print("SALES ANALYSIS COMPLETED")
print("Total Revenue:", total_revenue)
print("Best-selling Product:", best_product)
print(report)
print("Report saved as:", file_name)
