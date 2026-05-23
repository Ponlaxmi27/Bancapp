import pandas as pd

pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# ---------------------------------------------------
# 1. Read GL output file
# ---------------------------------------------------
gl_file = input("Enter GL output Excel path: ").strip()

gl_df = pd.read_excel(
    gl_file,
    sheet_name="Clean_Data",
    dtype=str
)

# ---------------------------------------------------
# 2. Read Transaction output file
# ---------------------------------------------------
transaction_file = input(
    "Enter Transaction output Excel path: "
).strip()

transaction_df = pd.read_excel(
    transaction_file,
    sheet_name="Clean_Data",
    dtype=str
)

# ---------------------------------------------------
# 3. Read CMS output file
# ---------------------------------------------------
cms_file = input(
    "Enter CMS output Excel path: "
).strip()

cms_df = pd.read_excel(
    cms_file,
    sheet_name="Clean_Data",
    dtype=str
)

# ---------------------------------------------------
# 4. Read Checkout output file
# ---------------------------------------------------
checkout_file = input(
    "Enter Checkout output Excel path: "
).strip()

checkout_df = pd.read_excel(
    checkout_file,
    sheet_name="Clean_Data",
    dtype=str
)

# ---------------------------------------------------
# 5. Clean join columns
# ---------------------------------------------------
gl_df["Description_First_14"] = (
    gl_df["Description_First_14"]
    .fillna("")
    .astype(str)
    .str.strip()
)

transaction_df["ContractNumber"] = (
    transaction_df["ContractNumber"]
    .fillna("")
    .astype(str)
    .str.strip()
)

cms_df["TARGET_NUMBER"] = (
    cms_df["TARGET_NUMBER"]
    .fillna("")
    .astype(str)
    .str.strip()
)

checkout_df["Reference"] = (
    checkout_df["Reference"]
    .fillna("")
    .astype(str)
    .str.strip()
)

# ---------------------------------------------------
# 6. LEFT JOIN GL + Transaction
# ---------------------------------------------------
joined_df = pd.merge(
    gl_df,
    transaction_df,
    how="left",
    left_on="ContractNumber",
    right_on="Description_First_14",
    suffixes=("_GL", "_Transaction")
)

# ---------------------------------------------------
# 7. LEFT JOIN with CMS
# ---------------------------------------------------
joined_df = pd.merge(
    joined_df,
    cms_df,
    how="left",
    left_on="ContractNumber",
    right_on="TARGET_NUMBER",
    suffixes=("", "_CMS")
)

# ---------------------------------------------------
# 8. LEFT JOIN with Checkout
# ---------------------------------------------------
joined_df = pd.merge(
    joined_df,
    checkout_df,
    how="left",
    left_on="WAY4Ref",
    right_on="Reference",
    suffixes=("", "_Checkout")
)

# ---------------------------------------------------
# 9. Match Status
# ---------------------------------------------------
joined_df["Transaction_Match"] = joined_df["Description_First_14"].apply(
    lambda x: "Matched" if pd.notna(x) and x != "" else "Not Matched"
)

joined_df["CMS_Match"] = joined_df["TARGET_NUMBER"].apply(
    lambda x: "Matched" if pd.notna(x) and x != "" else "Not Matched"
)

joined_df["Checkout_Match"] = joined_df["Reference"].apply(
    lambda x: "Matched" if pd.notna(x) and x != "" else "Not Matched"
)

# ---------------------------------------------------
# 10. Overall Match Status
# ---------------------------------------------------
joined_df["Overall_Status"] = joined_df.apply(
    lambda row:
    "Fully Matched"
    if (
            row["Transaction_Match"] == "Matched"
            and row["CMS_Match"] == "Matched"
            and row["Checkout_Match"] == "Matched"
    )
    else "Partially/Not Matched",
    axis=1
)

# ---------------------------------------------------
# 11. Summary
# ---------------------------------------------------
summary = pd.DataFrame({
    "Particulars": [
        "Total GL rows",
        "Transaction matched",
        "CMS matched",
        "Checkout matched",
        "Fully matched rows",
        "Partial/Not matched rows"
    ],
    "Value": [
        len(gl_df),

        len(joined_df[
                joined_df["Transaction_Match"] == "Matched"
                ]),

        len(joined_df[
                joined_df["CMS_Match"] == "Matched"
                ]),

        len(joined_df[
                joined_df["Checkout_Match"] == "Matched"
                ]),

        len(joined_df[
                joined_df["Overall_Status"] == "Fully Matched"
                ]),

        len(joined_df[
                joined_df["Overall_Status"]
                == "Partially/Not Matched"
                ])
    ]
})

# ---------------------------------------------------
# 12. Save output
# ---------------------------------------------------
output_file = "Final_Reconciliation_Output.xlsx"

with pd.ExcelWriter(output_file) as writer:

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    joined_df.to_excel(
        writer,
        sheet_name="Full_Left_Join",
        index=False
    )

    joined_df[
        joined_df["Overall_Status"] == "Fully Matched"
        ].to_excel(
        writer,
        sheet_name="Fully_Matched",
        index=False
    )

    joined_df[
        joined_df["Overall_Status"]
        == "Partially/Not Matched"
        ].to_excel(
        writer,
        sheet_name="Partial_Not_Matched",
        index=False
    )

# ---------------------------------------------------
# 13. Print result
# ---------------------------------------------------
print("Final LEFT JOIN reconciliation completed.")
print("Output file created:", output_file)

print("Total GL rows:", len(gl_df))

print(
    "Fully matched rows:",
    len(joined_df[
            joined_df["Overall_Status"] == "Fully Matched"
            ])
)

print(
    "Partial/Not matched rows:",
    len(joined_df[
            joined_df["Overall_Status"]
            == "Partially/Not Matched"
            ])
)