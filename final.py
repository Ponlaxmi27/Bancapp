import re
from pathlib import Path

import pandas as pd


# ============================================================
# 1. Input files
# ============================================================
# Keep the 4 input files in the same folder as this Python script.
# This avoids hardcoding any external folder path like Downloads.



gl_file = Path("/Users/ponlaxmi/Downloads/1121111_n.xlsx")
checkout_file = Path("/Users/ponlaxmi/Downloads/Checkout report_Online.xlsx")
transaction_file = Path("/Users/ponlaxmi/Downloads/CKO_transactions_2024-08-31.xlsx")
cms_file = Path("/Users/ponlaxmi/Downloads/CMS_Posted report.txt")


# ============================================================
# 2. Final output file
# ============================================================


output_file = Path("/Users/ponlaxmi/Downloads/final_join_using_script_clean_data.xlsx")


# ============================================================
# 3. Helper functions
# ============================================================

def clean_text(value):
    """Convert blank values to empty text and remove extra spaces."""
    if pd.isna(value):
        return ""
    return str(value).replace("\xa0", " ").strip()


def get_gl_reference(description):
    """Extract reference like 401-T-80475676 from GL Description."""
    description = clean_text(description)

    if description == "":
        return ""

    first_part = description.split("::")[0].strip()
    match = re.search(r"\d{3}-[A-Z]-\d+", first_part, flags=re.IGNORECASE)

    if match:
        return match.group(0).upper()

    return first_part.upper()


def make_common_data(df, report_name, reference_col, date_col, amount_col, recon_key_col):
    """
    Convert each report into same final columns:
    Report Name, Recon Key, Reference, Transaction Date, Amount
    """
    common = pd.DataFrame(index=df.index)
    common["Report Name"] = report_name
    common["Recon Key"] = df[recon_key_col]
    common["Reference"] = df[reference_col]
    common["Transaction Date"] = df[date_col]
    common["Amount"] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)
    return common


# ============================================================
# 4. Clean GL / Account data
# ============================================================

def clean_gl_data():
    # Header starts on Excel row 24, so Python uses header=23.
    df = pd.read_excel(gl_file, header=23)
    df.columns = [clean_text(col) for col in df.columns]

    df = df.dropna(how="all")
    df = df[df["Invoice Description"] != "Total"]

    df["Debit"] = pd.to_numeric(df["Debit"], errors="coerce").fillna(0)
    df["Credit"] = pd.to_numeric(df["Credit"], errors="coerce").fillna(0)
    df["GL Date"] = pd.to_datetime(df["GL Date"], errors="coerce")

    df["Description"] = df["Description"].apply(clean_text)
    df["Description First 14"] = df["Description"].str[:14]
    df["Reference"] = df["Description"].apply(get_gl_reference)

    # Amount = Debit - Credit
    df["Amount"] = df["Debit"] - df["Credit"]

    # Duplicate removal rule:
    # Account + first 14 characters of Description + Credit + GL Date
    df = df.drop_duplicates(
        subset=["Account", "Description First 14", "Credit", "GL Date"],
        keep="first",
    )

    df["Transaction Date"] = df["GL Date"].dt.strftime("%Y-%m-%d")

    df["Recon Key"] = (
            df["Account"].astype(str)
            + "|"
            + df["Description First 14"].astype(str)
            + "|"
            + df["Credit"].astype(str)
            + "|"
            + df["Transaction Date"].astype(str)
    )

    return df


# ============================================================
# 5. Clean Checkout data
# ============================================================

def clean_checkout_data():
    df = pd.read_excel(
        checkout_file,
        sheet_name="in",
        dtype={"Reference": str, "Payment ID": str},
    )

    df.columns = df.columns.str.strip()

    df["Reference"] = df["Reference"].fillna("").astype(str).str.strip()
    df["Payment ID"] = df["Payment ID"].fillna("").astype(str).str.strip()
    df["Action Type"] = df["Action Type"].astype(str).str.strip()
    df["Response Description"] = df["Response Description"].astype(str).str.strip()

    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    df["Action Date UTC"] = pd.to_datetime(df["Action Date UTC"], errors="coerce")
    df["Action Date"] = df["Action Date UTC"].dt.strftime("%Y-%m-%d")

    df["Is Approved"] = df["Response Description"].eq("Approved")

    # Settlement Amount:
    # Approved Capture = positive
    # Approved Refund = negative
    # Other rows = zero
    df["Settlement Amount"] = 0.0

    df.loc[
        (df["Action Type"] == "Capture") & (df["Is Approved"]),
        "Settlement Amount",
    ] = df["Amount"]

    df.loc[
        (df["Action Type"] == "Refund") & (df["Is Approved"]),
        "Settlement Amount",
    ] = -df["Amount"]

    # Recon Key = Reference + Payment ID + latest action date
    df["Latest Action Date"] = (
        df.groupby(["Reference", "Payment ID"])["Action Date UTC"]
        .transform("max")
        .dt.strftime("%Y-%m-%d")
    )

    df["Recon Key"] = (
            df["Reference"]
            + "|"
            + df["Payment ID"]
            + "|"
            + df["Latest Action Date"].fillna("")
    )

    return df


# ============================================================
# 6. Clean CKO transaction data
# ============================================================

def clean_transaction_data():
    df = pd.read_excel(
        transaction_file,
        sheet_name="CKO Transactions",
        dtype={
            "ContractNumber": str,
            "RRNAuth": str,
            "RRNCapt": str,
            "WAY4Ref": str,
        },
    )

    df.columns = df.columns.str.strip()

    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    df["Created"] = pd.to_datetime(df["Created"], errors="coerce")
    df["Transaction Date"] = df["Created"].dt.strftime("%Y-%m-%d")

    df["ContractNumber"] = df["ContractNumber"].fillna("").astype(str).str.strip()
    df["RRNCapt"] = df["RRNCapt"].fillna("").astype(str).str.strip()
    df["WAY4Ref"] = df["WAY4Ref"].fillna("").astype(str).str.strip()

    df["Recon Key"] = (
            df["ContractNumber"]
            + "|"
            + df["RRNCapt"]
            + "|"
            + df["WAY4Ref"]
    )

    return df


# ============================================================
# 7. Clean CMS posted data
# ============================================================

def clean_cms_data():
    df = pd.read_csv(cms_file, sep="\t", dtype=str)
    df.columns = df.columns.str.strip()

    df["TARGET_NUMBER"] = df["TARGET_NUMBER"].fillna("").astype(str).str.strip()
    df["SETTL_AMOUNT"] = df["SETTL_AMOUNT"].fillna("0").astype(str).str.strip()
    df["TRANS_TYPE"] = df["TRANS_TYPE"].fillna("").astype(str).str.strip()

    df["SETTL_AMOUNT"] = (
        df["SETTL_AMOUNT"]
        .str.replace(",", "", regex=False)
        .str.replace('"', "", regex=False)
    )

    df["SETTL_AMOUNT"] = pd.to_numeric(df["SETTL_AMOUNT"], errors="coerce").fillna(0)

    df["TRANS_DATE"] = pd.to_datetime(
        df["TRANS_DATE"],
        format="%d/%m/%y %H:%M",
        errors="coerce",
    )

    df["Transaction Date"] = df["TRANS_DATE"].dt.strftime("%Y-%m-%d")
    df["Recon Key"] = df["TARGET_NUMBER"]

    return df


# ============================================================
# 8. Run all cleaning steps
# ============================================================

gl_clean = clean_gl_data()
checkout_clean = clean_checkout_data()
transaction_clean = clean_transaction_data()
cms_clean = clean_cms_data()


# ============================================================
# 9. Convert all clean data into common format
# ============================================================

gl_common = make_common_data(
    gl_clean,
    "GL / Account",
    "Reference",
    "Transaction Date",
    "Amount",
    "Recon Key",
)

checkout_common = make_common_data(
    checkout_clean,
    "Checkout",
    "Reference",
    "Action Date",
    "Settlement Amount",
    "Recon Key",
)

transaction_common = make_common_data(
    transaction_clean,
    "CKO Transaction",
    "ContractNumber",
    "Transaction Date",
    "Amount",
    "Recon Key",
)

cms_common = make_common_data(
    cms_clean,
    "CMS Posted",
    "TARGET_NUMBER",
    "Transaction Date",
    "SETTL_AMOUNT",
    "Recon Key",
)


# ============================================================
# 10. Join all clean data
# ============================================================

all_clean_data = pd.concat(
    [gl_common, checkout_common, transaction_common, cms_common],
    ignore_index=True,
)

# ============================================================
# Match status across all combinations
# ============================================================

# Your all_clean_data should have these columns:
# Report Name
# Reference
# Amount

report_name_map = {
    "GL / Account": "GL",
    "Checkout": "Checkout",
    "CKO Transaction": "Transactions",
    "CMS Posted": "CMS",
}

all_clean_data["Report Group"] = all_clean_data["Report Name"].map(report_name_map)

# We compare using Reference
all_clean_data["Match Key"] = (
    all_clean_data["Reference"]
    .fillna("")
    .astype(str)
    .str.strip()
)

reports_to_check = ["CMS", "GL", "Transactions", "Checkout"]

# Create availability table
presence = (
    all_clean_data.assign(Available=1)
    .pivot_table(
        index="Match Key",
        columns="Report Group",
        values="Available",
        aggfunc="max",
        fill_value=0,
    )
    .reset_index()
)

# Make sure all report columns exist
for report in reports_to_check:
    if report not in presence.columns:
        presence[report] = 0


# Create amount table
amount_summary = (
    all_clean_data
    .groupby(["Match Key", "Report Group"])["Amount"]
    .sum()
    .unstack(fill_value=0)
    .reset_index()
)

for report in reports_to_check:
    if report not in amount_summary.columns:
        amount_summary[report] = 0

amount_summary = amount_summary.rename(
    columns={
        "CMS": "CMS Amount",
        "GL": "GL Amount",
        "Transactions": "Transactions Amount",
        "Checkout": "Checkout Amount",
    }
)


# Join availability and amount
match_status = presence.merge(amount_summary, on="Match Key", how="left")

# ============================================================
# Pair-wise Match Summary
# ============================================================

pair_list = [
    ("GL", "CMS"),
    ("GL", "Transactions"),
    ("GL", "Checkout"),
    ("CMS", "Transactions"),
    ("CMS", "Checkout"),
    ("Transactions", "Checkout"),
]

pair_summary_rows = []

for report_1, report_2 in pair_list:
    matched_count = len(
        match_status[
            (match_status[report_1] == 1)
            & (match_status[report_2] == 1)
            ]
    )

    only_report_1_count = len(
        match_status[
            (match_status[report_1] == 1)
            & (match_status[report_2] == 0)
            ]
    )

    only_report_2_count = len(
        match_status[
            (match_status[report_1] == 0)
            & (match_status[report_2] == 1)
            ]
    )

    pair_summary_rows.append({
        "Report Combination": report_1 + " vs " + report_2,
        "Matched in Both": matched_count,
        "Available only in " + report_1: only_report_1_count,
        "Available only in " + report_2: only_report_2_count,
    })

pair_wise_summary = pd.DataFrame(pair_summary_rows)

# Function to describe match status
def describe_match(row):
    available = []
    missing = []

    for report in reports_to_check:
        if row[report] == 1:
            available.append(report)
        else:
            missing.append(report)

    if len(available) == len(reports_to_check):
        return "Matched in All Reports"

    if len(available) == 1:
        return "Available only in " + available[0]

    return (
            "Available in "
            + " and ".join(available)
            + " but Not Available in "
            + " and ".join(missing)
    )


match_status["Match Status"] = match_status.apply(describe_match, axis=1)

# Count how many reports contain the reference
match_status["Available Count"] = match_status[reports_to_check].sum(axis=1)



# Separate outputs
matched_all = match_status[match_status["Available Count"] == 4]

partial_matches = match_status[
    (match_status["Available Count"] > 1)
    & (match_status["Available Count"] < 4)
    ]

completely_not_matched = match_status[
    match_status["Available Count"] == 1
    ]

# ============================================================
# Match Summary
# ============================================================

match_summary = pd.DataFrame({
    "Particulars": [
        "Matched in All Reports",
        "Partial Matches",
        "Completely Not Matched",
        "Total Unique References"
    ],
    "Count": [
        len(matched_all),
        len(partial_matches),
        len(completely_not_matched),
        len(match_status)
    ]
})
# ============================================================
# 11. Match status across reports
# ============================================================
# This section checks whether each Reference is available in:
# GL, Checkout, Transactions, and CMS.

report_name_map = {
    "GL / Account": "GL",
    "Checkout": "Checkout",
    "CKO Transaction": "Transactions",
    "CMS Posted": "CMS",
}

all_clean_data["Report Group"] = all_clean_data["Report Name"].map(report_name_map)
all_clean_data["Match Key"] = all_clean_data["Reference"].fillna("").astype(str).str.strip()

reports_to_check = ["GL", "Checkout", "Transactions", "CMS"]

presence = (
    all_clean_data.assign(Available=1)
    .pivot_table(
        index="Match Key",
        columns="Report Group",
        values="Available",
        aggfunc="max",
        fill_value=0,
    )
    .reset_index()
)

for report in reports_to_check:
    if report not in presence.columns:
        presence[report] = 0

amount_summary = (
    all_clean_data.groupby(["Match Key", "Report Group"])["Amount"]
    .sum()
    .unstack(fill_value=0)
    .reset_index()
)

for report in reports_to_check:
    if report not in amount_summary.columns:
        amount_summary[report] = 0

amount_summary = amount_summary.rename(
    columns={report: f"{report} Amount" for report in reports_to_check}
)

match_status = presence.merge(amount_summary, on="Match Key", how="left")


def describe_match(row):
    available = [report for report in reports_to_check if row[report] == 1]
    missing = [report for report in reports_to_check if row[report] == 0]

    if len(available) == len(reports_to_check):
        return "Matched in All Reports"

    if len(available) == 1:
        return f"Completely Not Matched - Available only in {available[0]}"

    available_text = " and ".join(available)
    missing_text = " and ".join(missing)
    return f"Partial Match - Available in {available_text} but Not Available in {missing_text}"


match_status["Match Status"] = match_status.apply(describe_match, axis=1)
match_status["Available Count"] = match_status[reports_to_check].sum(axis=1)

matched_all = match_status[match_status["Available Count"] == len(reports_to_check)]
completely_not_matched = match_status[match_status["Available Count"] == 1]
partial_matches = match_status[
    (match_status["Available Count"] > 1)
    & (match_status["Available Count"] < len(reports_to_check))
    ]

# ============================================================
# Final Match Summary With Report Names
# ============================================================

final_match_summary = pd.DataFrame({

})

report_wise_match_summary = pd.DataFrame({
    "Report Name": ["CMS", "GL", "Transactions", "Checkout"],
    "Available Count": [
        match_status["CMS"].sum(),
        match_status["GL"].sum(),
        match_status["Transactions"].sum(),
        match_status["Checkout"].sum()
    ],
    "Not Available Count": [
        len(match_status) - match_status["CMS"].sum(),
        len(match_status) - match_status["GL"].sum(),
        len(match_status) - match_status["Transactions"].sum(),
        len(match_status) - match_status["Checkout"].sum()
    ]
})
# ============================================================
# 12. Create summary
# ============================================================

summary = (
    all_clean_data.groupby("Report Name")
    .agg(
        Rows=("Amount", "count"),
        Total_Amount=("Amount", "sum"),
        Unique_Recon_Keys=("Recon Key", "nunique"),
    )
    .reset_index()
)

grand_total = pd.DataFrame({
    "Report Name": ["Grand Total"],
    "Rows": [len(all_clean_data)],
    "Total_Amount": [all_clean_data["Amount"].sum()],
    "Unique_Recon_Keys": [all_clean_data["Recon Key"].nunique()],
})

summary = pd.concat([summary, grand_total], ignore_index=True)


# ============================================================
# 13. Save final output
# ============================================================
with pd.ExcelWriter(output_file) as writer:
    summary.to_excel(writer, sheet_name="Summary", index=False, startrow=0)

    final_match_summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False,
        startrow=len(summary) + 3
    )


    pair_wise_summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False,
        startrow=len(summary)  + 9
    )

    match_status.to_excel(writer, sheet_name="Match_Status_All", index=False)
    matched_all.to_excel(writer, sheet_name="Matched_All", index=False)
    partial_matches.to_excel(writer, sheet_name="Partial_Matches", index=False)
    completely_not_matched.to_excel(writer, sheet_name="Completely_Not_Matched", index=False)

    all_clean_data.to_excel(writer, sheet_name="All_Clean_Data", index=False)


# ============================================================
# 14. Print result
# ============================================================

print("Final joined file created using clean data from this script.")
print("Output file:", output_file)
print("")
print(summary.to_string(index=False))
print("")
print("Matched in all reports:", len(matched_all))
print("Partial matches:", len(partial_matches))
print("Completely not matched:", len(completely_not_matched))
