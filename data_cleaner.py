
import pandas as pd

# Create sample messy business data
data = {
    "Name": [" John ", "Mary", "Peter ", "John", None, "Sarah"],
    "Department": [" Sales", "IT ", "Finance", "Sales", "IT", None],
    "Salary": [50000, 65000, None, 50000, 70000, 60000],
    "Email": [
        "JOHN@email.com",
        "mary@email.com",
        "peter@email.com",
        "JOHN@email.com",
        "test@email.com",
        "sarah@email.com"
    ]
}

df = pd.DataFrame(data)

print("Original Data:")
print(df)

# Remove extra spaces
for column in ["Name", "Department", "Email"]:
    df[column] = df[column].str.strip()

# Standardise email addresses
df["Email"] = df["Email"].str.lower()

# Remove duplicate records
df = df.drop_duplicates()

# Fill missing salaries with average salary
df["Salary"] = df["Salary"].fillna(df["Salary"].mean())

# Identify incomplete records
df["Name"] = df["Name"].fillna("Unknown")
df["Department"] = df["Department"].fillna("Unassigned")

# Save cleaned data
df.to_excel("Cleaned_Business_Data.xlsx", index=False)

print("\nCleaned Data:")
print(df)

print("\nData cleaning completed successfully!")
print("Saved as Cleaned_Business_Data.xlsx")
