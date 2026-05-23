import re
from pathlib import Path
import pandas as pd
import numpy as np

#pd.set_option('display.max_rows', None)      # Show all rows
#pd.set_option('display.max_columns', None)   # Show all columns
pd.set_option('display.width', None)         # Full width
pd.set_option('display.max_colwidth', None)  # Full column content

input_file = Path(input("Enter Excel file path: ").strip())
output_file = "/Users/ponlaxmi/Downloads/Checkout_output.xlsx"

# 2. Read the Excel file
# dtype keeps big reference numbers as text, so Python does not change them to scientific notation.
df = pd.read_excel(input_file, sheet_name="in", dtype={"Reference": str, "Payment ID": str})


# 3. Clean column names and text columns
df.columns = df.columns.str.strip()

df["Action Type"] = df["Action Type"].astype(str).str.strip()
df["Response Description"] = df["Response Description"].astype(str).str.strip()
df["Payment Method Name"] = df["Payment Method Name"].astype(str).str.strip()
df["Currency Symbol"] = df["Currency Symbol"].astype(str).str.strip()
df["Reference"] = df["Reference"].fillna("").astype(str).str.strip()
df["Payment ID"] = df["Payment ID"].fillna("").astype(str).str.strip()


# 4. Convert Amount column into number
df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)


# 5. Convert Action Date UTC column into date/time
# If you get an error here, first check the exact column name in your Excel file.
date_column = "Action Date UTC"

if date_column not in df.columns:
    print("Date column not found.")
    print("Available columns are:")
    print(df.columns.tolist())
    raise SystemExit

df[date_column] = pd.to_datetime(df[date_column], errors="coerce")


# 6. Create only date column for daily summary
# Example output: 2024-08-31
df["Action Date"] = df[date_column].dt.strftime("%Y-%m-%d")


# 7. Identify approved transactions
# In this report, Approved usually has Response Code 10000.
df["Is Approved"] = df["Response Description"].eq("Approved")

# 8. Create settlement amount
# Capture = money received
# Refund = money returned
# Other action types, like Authorisation, are not settlement money
df["Settlement Amount"] = 0.0

df.loc[
    (df["Action Type"] == "Capture") & (df["Is Approved"]),
    "Settlement Amount"
] = df["Amount"]

df.loc[
    (df["Action Type"] == "Refund") & (df["Is Approved"]),
    "Settlement Amount"
] = -df["Amount"]


# 9. Create a reconciliation key
# Recon Key = Reference + Payment ID + Latest Action Date
#
# First, find the latest action date for each Reference and Payment ID.
# transform("max") puts that latest date back on every row in the same group.
df["Latest Action Date"] = (
    df.groupby(["Reference", "Payment ID"])[date_column]
    .transform("max")
    .dt.strftime("%Y-%m-%d")
)

# Now create one combined key.
df["Recon Key"] = (
        df["Reference"]
        + "|"
        + df["Payment ID"]
        + "|"
        + df["Latest Action Date"].fillna("")
)


# 10. Summary by payment
payment_reconciliation = (
    df.groupby("Recon Key", dropna=False)
    .agg(
        Reference=("Reference", "first"),
        Payment_ID=("Payment ID", "first"),
        Latest_Action_Date=("Latest Action Date", "first"),
        Currency=("Currency Symbol", "first"),
        Capture_Amount=("Settlement Amount", lambda x: x[x > 0].sum()),
        Refund_Amount=("Settlement Amount", lambda x: -x[x < 0].sum()),
        Net_Amount=("Settlement Amount", "sum"),
        Total_Rows=("Action ID", "count"),
        First_Date=(date_column, "min"),
        Last_Date=(date_column, "max"),
    )
    .reset_index()
)


# 11. Find unmatched items
# Unmatched Capture = captured but no refund
# Unmatched Refund = refunded but no capture
unmatched_captures = payment_reconciliation[
    (payment_reconciliation["Capture_Amount"] > 0)
    & (payment_reconciliation["Refund_Amount"] == 0)
    ]

unmatched_refunds = payment_reconciliation[
    (payment_reconciliation["Refund_Amount"] > 0)
    & (payment_reconciliation["Capture_Amount"] == 0)
    ]


# 12. Summary by action type
by_action_type = (
    df.groupby(["Action Type", "Response Description"], dropna=False)
    .agg(
        Rows=("Action ID", "count"),
        Amount=("Amount", "sum"),
        Settlement_Amount=("Settlement Amount", "sum"),
    )
    .reset_index()
)


# 13. Summary by payment method
by_payment_method = (
    df.groupby("Payment Method Name", dropna=False)
    .agg(
        Rows=("Action ID", "count"),
        Amount=("Amount", "sum"),
        Settlement_Amount=("Settlement Amount", "sum"),
    )
    .reset_index()
)


# 14. Daily summary
daily_summary = (
    df.groupby("Action Date", dropna=False)
    .agg(
        Rows=("Action ID", "count"),
        Capture_Amount=("Settlement Amount", lambda x: x[x > 0].sum()),
        Refund_Amount=("Settlement Amount", lambda x: -x[x < 0].sum()),
        Net_Amount=("Settlement Amount", "sum"),
    )
    .reset_index()
)


# 15. Main summary
total_capture = df.loc[
    (df["Action Type"] == "Capture") & (df["Is Approved"]),
    "Amount"
].sum()

total_refund = df.loc[
    (df["Action Type"] == "Refund") & (df["Is Approved"]),
    "Amount"
].sum()

summary = pd.DataFrame({
    "Particulars": [
        "Total rows",
        "Approved capture amount",
        "Approved refund amount",
        "Net settlement amount",
        "Unmatched capture payments",
        "Unmatched refund payments",
    ],
    "Value": [
        len(df),
        total_capture,
        total_refund,
        total_capture - total_refund,
        len(unmatched_captures),
        len(unmatched_refunds),
        ],
})


# 16. Save everything into one Excel output file
with pd.ExcelWriter(output_file) as writer:
    summary.to_excel(writer, sheet_name="Summary", index=False)
    payment_reconciliation.to_excel(writer, sheet_name="Payment_Recon", index=False)
    unmatched_captures.to_excel(writer, sheet_name="Unmatched_Captures", index=False)
    unmatched_refunds.to_excel(writer, sheet_name="Unmatched_Refunds", index=False)
    by_action_type.to_excel(writer, sheet_name="By_Action_Type", index=False)
    by_payment_method.to_excel(writer, sheet_name="By_Payment_Method", index=False)
    daily_summary.to_excel(writer, sheet_name="Daily_Summary", index=False)
    df.to_excel(writer, sheet_name="Clean_Data", index=False)


print("Checkout reconciliation completed.")
print("Output file created:", output_file)
print("Approved capture amount:", total_capture)
print("Approved refund amount:", total_refund)
print("Net settlement amount:", total_capture - total_refund)