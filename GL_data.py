import re
from pathlib import Path
import pandas as pd
import numpy as np

#pd.set_option('display.max_rows', None)      # Show all rows
#pd.set_option('display.max_columns', None)   # Show all columns
pd.set_option('display.width', None)         # Full width
pd.set_option('display.max_colwidth', None)  # Full column content

input_file = Path(input("Enter Excel file path: ").strip())
output_file = "/Users/ponlaxmi/Downloads/reconciliation_output.xlsx"

if not input_file.exists():
    print("File not found:", input_file)
    raise SystemExit()


# Read Excel file
# Your actual table header starts on row 24
# Python counts from 0, so we use header=23
df = pd.read_excel(input_file, header=23)


# Remove completely blank rows
df = df.dropna(how="all")


# Remove the final Total row
df = df[df["Invoice Description"] != "Total"]


# Convert Credit column to number
df["Credit"] = pd.to_numeric(df["Credit"], errors="coerce").fillna(0)


# Convert GL Date column to date
df["GL Date"] = pd.to_datetime(df["GL Date"], errors="coerce")


# Convert Description to text
df["Description"] = df["Description"].astype(str)


# Get first 14 characters from Description
df["Description_First_14"] = df["Description"].str[:14]

# Group Source column
def source_group(source):
    source = str(source).strip()

    if source in ["TASHEEL", "TAQSEET2", "AutoCopy"]:
        return "Auto"
    elif source in ["Manual", "Spreadsheet"]:
        return "Manual"
    else:
        return source


df["Source_Group"] = df["Source."].apply(source_group)

# Remove duplicates
# Duplicate means same Account, same first 14 characters of Description,
# same Credit, and same GL Date
df_clean = df.drop_duplicates(
    subset=["Account", "Description_First_14", "Credit", "GL Date"],
    keep="first"
)


# Create Source summary
source_summary = df_clean.groupby("Source_Group").agg(
    Rows=("Account", "size"),
    Debit=("Debit", "sum"),
    Credit=("Credit", "sum")
).reset_index()

source_summary["Net"] = source_summary["Debit"] - source_summary["Credit"]


# Create main summary
main_summary = pd.DataFrame({
    "Particulars": [
        "Rows before duplicate removal",
        "Rows after duplicate removal",
        "Duplicate rows removed",
        "Total Debit after duplicate removal",
        "Total Credit after duplicate removal",
        "Net Balance after duplicate removal"
    ],
    "Amount": [
        len(df),
        len(df_clean),
        len(df) - len(df_clean),
        df_clean["Debit"].sum(),
        df_clean["Credit"].sum(),
        df_clean["Debit"].sum() - df_clean["Credit"].sum()
    ]
})


# Save output to Excel
with pd.ExcelWriter(output_file) as writer:
    main_summary.to_excel(writer, sheet_name="Summary", index=False)
    source_summary.to_excel(writer, sheet_name="By_Source", index=False)
    df_clean.to_excel(writer, sheet_name="Clean_Data", index=False)


print("Reconciliation completed")
print("Output file created:", output_file)

