
import pandas as pd
import matplotlib.pyplot as plt

# Load business intelligence data
file_name = "Business_Intelligence_Report.xlsx"

df = pd.read_excel(file_name, sheet_name="Transactions")

# Convert date column
df["Date"] = pd.to_datetime(df["Date"])

# Calculate monthly performance
monthly = df.groupby(
    [df["Date"].dt.to_period("M"), "Type"]
)["Amount"].sum().unstack(fill_value=0)

monthly["Profit"] = monthly.get("Income", 0) - monthly.get("Expense", 0)

# Display business performance
print("\n===== BUSINESS PERFORMANCE DASHBOARD =====")
print(monthly)

total_income = df[df["Type"] == "Income"]["Amount"].sum()
total_expenses = df[df["Type"] == "Expense"]["Amount"].sum()
total_profit = total_income - total_expenses

print("\nTotal Income:", total_income)
print("Total Expenses:", total_expenses)
print("Total Profit:", total_profit)

if total_income > 0:
    profit_margin = (total_profit / total_income) * 100
else:
    profit_margin = 0

print("Profit Margin:", round(profit_margin, 2), "%")

# Create charts
monthly[["Income", "Expense"]].plot(
    kind="bar",
    title="Monthly Income vs Expenses"
)

plt.ylabel("Amount")
plt.tight_layout()
plt.savefig("monthly_performance.png")
plt.show()

monthly["Profit"].plot(
    kind="line",
    marker="o",
    title="Monthly Profit Trend"
)

plt.ylabel("Profit")
plt.tight_layout()
plt.savefig("profit_trend.png")
plt.show()

# Export analysis
monthly.to_excel("Advanced_Business_Dashboard.xlsx")

print("\nDashboard completed successfully!")
