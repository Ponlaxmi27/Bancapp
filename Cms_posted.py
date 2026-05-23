import pandas as pd
from pathlib import Path

pd.set_option('display.width', None)         # Full width
pd.set_option('display.max_colwidth', None)  # Full column content

input_file = Path(input("Enter Excel file path: ").strip())
output_file = "/Users/ponlaxmi/Downloads/Cms_reconciliation_output.xlsx"


# 2. Read the text file
# This file is tab-separated, so we use sep="\t"
df = pd.read_csv(input_file, sep="\t", dtype=str)


# 3. Clean column names
df.columns = df.columns.str.strip()


# 4. Clean text columns
df["TARGET_NUMBER"] = df["TARGET_NUMBER"].fillna("").astype(str).str.strip()
df["SETTL_AMOUNT"] = df["SETTL_AMOUNT"].fillna("0").astype(str).str.strip()
df["TRANS_DATE"] = df["TRANS_DATE"].fillna("").astype(str).str.strip()
df["TRANS_TYPE"] = df["TRANS_TYPE"].fillna("").astype(str).str.strip()


# 5. Convert settlement amount into number
# Some amounts are like "1,146.00", so we remove comma first.
df["SETTL_AMOUNT"] = (
    df["SETTL_AMOUNT"]
    .str.replace(",", "", regex=False)
    .str.replace('"', "", regex=False)
)

df["SETTL_AMOUNT"] = pd.to_numeric(df["SETTL_AMOUNT"], errors="coerce").fillna(0)


# 6. Convert transaction date into proper date/time
# Format in file looks like: 23/08/24 2:23
# This means: day/month/year hour:minute
df["TRANS_DATE"] = pd.to_datetime(df["TRANS_DATE"], format="%d/%m/%y %H:%M", errors="coerce")


# 7. Create date columns in clean format
df["Transaction DateTime"] = df["TRANS_DATE"].dt.strftime("%Y-%m-%d %H:%M:%S")
df["Transaction Date"] = df["TRANS_DATE"].dt.strftime("%Y-%m-%d")


# 8. Create reconciliation key
# Here TARGET_NUMBER is the main reference number.
df["Recon Key"] = df["TARGET_NUMBER"]

# 9. Find duplicate target numbers
duplicate_targets = df[df.duplicated(subset=["TARGET_NUMBER"], keep=False)]


# 10. Summary by transaction type
by_transaction_type = (
    df.groupby("TRANS_TYPE")
    .agg(
        Rows=("TARGET_NUMBER", "count"),
        Total_Amount=("SETTL_AMOUNT", "sum"),
    )
    .reset_index()
)


# 11. Daily summary
daily_summary = (
    df.groupby("Transaction Date")
    .agg(
        Rows=("TARGET_NUMBER", "count"),
        Total_Amount=("SETTL_AMOUNT", "sum"),
    )
    .reset_index()
)


# 12. Main summary
summary = pd.DataFrame({
    "Particulars": [
        "Total rows",
        "Total settlement amount",
        "Unique target numbers",
        "Duplicate target rows",
        "Earliest transaction date",
        "Latest transaction date",
    ],
    "Value": [
        len(df),
        df["SETTL_AMOUNT"].sum(),
        df["TARGET_NUMBER"].nunique(),
        len(duplicate_targets),
        df["Transaction DateTime"].min(),
        df["Transaction DateTime"].max(),
    ],
})


# 13. Save reconciliation output into Excel
with pd.ExcelWriter(output_file) as writer:
    summary.to_excel(writer, sheet_name="Summary", index=False)
    by_transaction_type.to_excel(writer, sheet_name="By_Transaction_Type", index=False)
    daily_summary.to_excel(writer, sheet_name="Daily_Summary", index=False)
    duplicate_targets.to_excel(writer, sheet_name="Duplicate_Targets", index=False)
    df.to_excel(writer, sheet_name="Clean_Data", index=False)


# 14. Print result
print("CMS posted reconciliation completed.")
print("Output file created:", output_file)
print("Total rows:", len(df))
print("Total settlement amount:", df["SETTL_AMOUNT"].sum())
print("Duplicate target rows:", len(duplicate_targets))