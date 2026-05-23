import pandas as pd
from pathlib import Path

pd.set_option('display.width', None)         # Full width
pd.set_option('display.max_colwidth', None)  # Full column content

input_file = Path(input("Enter Excel file path: ").strip())
output_file = "/Users/ponlaxmi/Downloads/Transaction_reconciliation_output.xlsx"


# 2. Read the Excel file
df = pd.read_excel(
    input_file,
    sheet_name="CKO Transactions",
    dtype={
        "ContractNumber": str,
        "RRNAuth": str,
        "RRNCapt": str,
        "WAY4Ref": str,
    },
)


# 3. Clean column names
df.columns = df.columns.str.strip()


# 4. Convert Amount into number
df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)


# 5. Convert Created into date/time
df["Created"] = pd.to_datetime(df["Created"], errors="coerce")


# 6. Create date columns in proper format
# Created DateTime example: 2024-08-29 16:49:28
# Created Date example:     2024-08-29
df["Created DateTime"] = df["Created"].dt.strftime("%Y-%m-%d %H:%M:%S")
df["Created Date"] = df["Created"].dt.strftime("%Y-%m-%d")


# 7. Clean text columns
df["ContractNumber"] = df["ContractNumber"].astype(str).str.strip()
df["RRNAuth"] = df["RRNAuth"].astype(str).str.strip()
df["RRNCapt"] = df["RRNCapt"].astype(str).str.strip()
df["Payment Type"] = df["Payment Type"].astype(str).str.strip()
df["WAY4Ref"] = df["WAY4Ref"].fillna("").astype(str).str.strip()


# 8. Convert IsWAY4Posted into number
# 1 means posted to WAY4
# 0 means not posted to WAY4
df["IsWAY4Posted"] = pd.to_numeric(df["IsWAY4Posted"], errors="coerce").fillna(0)


# 9. Split data into posted and not posted
posted = df[df["IsWAY4Posted"] == 1]
not_posted = df[df["IsWAY4Posted"] == 0]


# 10. Find duplicate checks
# RRNAuth and RRNCapt should normally be unique transaction references.
duplicate_rrn_auth = df[df.duplicated(subset=["RRNAuth"], keep=False)]
duplicate_rrn_capt = df[df.duplicated(subset=["RRNCapt"], keep=False)]
duplicate_way4_ref = posted[posted.duplicated(subset=["WAY4Ref"], keep=False)]


# 11. Summary by posting status
by_posting_status = (
    df.groupby("IsWAY4Posted")
    .agg(
        Rows=("Amount", "count"),
        Total_Amount=("Amount", "sum"),
    )
    .reset_index()
)


# 12. Summary by payment type
by_payment_type = (
    df.groupby("Payment Type")
    .agg(
        Rows=("Amount", "count"),
        Total_Amount=("Amount", "sum"),
    )
    .reset_index()
)


# 13. Daily summary
daily_summary = (
    df.groupby("Created Date")
    .agg(
        Rows=("Amount", "count"),
        Total_Amount=("Amount", "sum"),
    )
    .reset_index()
)


# 14. Main summary
summary = pd.DataFrame({
    "Particulars": [
        "Total rows",
        "Total amount",
        "Posted rows",
        "Posted amount",
        "Not posted rows",
        "Not posted amount",
        "Duplicate RRNAuth rows",
        "Duplicate RRNCapt rows",
        "Duplicate WAY4Ref rows",
    ],
    "Value": [
        len(df),
        df["Amount"].sum(),
        len(posted),
        posted["Amount"].sum(),
        len(not_posted),
        not_posted["Amount"].sum(),
        len(duplicate_rrn_auth),
        len(duplicate_rrn_capt),
        len(duplicate_way4_ref),
    ],
})


# 15. Save the reconciliation into Excel
with pd.ExcelWriter(output_file) as writer:
    summary.to_excel(writer, sheet_name="Summary", index=False)
    by_posting_status.to_excel(writer, sheet_name="By_Posting_Status", index=False)
    by_payment_type.to_excel(writer, sheet_name="By_Payment_Type", index=False)
    daily_summary.to_excel(writer, sheet_name="Daily_Summary", index=False)
    not_posted.to_excel(writer, sheet_name="Not_Posted", index=False)
    posted.to_excel(writer, sheet_name="Posted", index=False)
    duplicate_rrn_auth.to_excel(writer, sheet_name="Duplicate_RRNAuth", index=False)
    duplicate_rrn_capt.to_excel(writer, sheet_name="Duplicate_RRNCapt", index=False)
    duplicate_way4_ref.to_excel(writer, sheet_name="Duplicate_WAY4Ref", index=False)
    df.to_excel(writer, sheet_name="Clean_Data", index=False)


# 16. Print result
print("CKO reconciliation completed.")
print("Output file created:", output_file)
print("Total amount:", df["Amount"].sum())
print("Posted amount:", posted["Amount"].sum())
print("Not posted amount:", not_posted["Amount"].sum())