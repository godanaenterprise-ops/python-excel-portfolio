
import pandas as pd
from datetime import datetime
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

print("EXCEL FINANCIAL REPORT GENERATOR")
print("--------------------------------")

transactions = []

while True:
    category = input(
        "\nEnter category (or type done): "
    )

    if category.lower() == "done":
        break

    transaction_type = input(
        "Enter type (Income/Expense): "
    ).capitalize()

    if transaction_type not in ["Income", "Expense"]:
        print("Please enter Income or Expense.")
        continue

    try:
        amount = float(input("Enter amount in KSh: "))

        if amount <= 0:
            print("Amount must be greater than zero.")
            continue

        transactions.append({
            "Date": datetime.now().strftime("%Y-%m-%d"),
            "Category": category,
            "Type": transaction_type,
            "Amount": amount
        })

    except ValueError:
        print("Please enter a valid number.")

if not transactions:
    print("No records entered. Exiting.")
    raise SystemExit

df = pd.DataFrame(transactions)

income = df.loc[
    df["Type"] == "Income", "Amount"
].sum()

expenses = df.loc[
    df["Type"] == "Expense", "Amount"
].sum()

profit = income - expenses

summary = pd.DataFrame({
    "Metric": ["Total Income", "Total Expenses", "Net Profit"],
    "Amount (KSh)": [income, expenses, profit]
})

filename = "Financial_Report.xlsx"

with pd.ExcelWriter(filename, engine="openpyxl") as writer:
    df.to_excel(
        writer,
        sheet_name="Transactions",
        index=False
    )

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    workbook = writer.book

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for cell in worksheet[1]:
            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )
            cell.fill = PatternFill(
                start_color="1F4E78",
                end_color="1F4E78",
                fill_type="solid"
            )

        for column in worksheet.columns:
            max_length = max(
                len(str(cell.value or ""))
                for cell in column
            )
            worksheet.column_dimensions[
                get_column_letter(column[0].column)
            ].width = max_length + 4

print("\nREPORT GENERATED SUCCESSFULLY!")
print(f"Income: KSh {income:,.2f}")
print(f"Expenses: KSh {expenses:,.2f}")
print(f"Profit: KSh {profit:,.2f}")
print(f"Excel file: {filename}")
