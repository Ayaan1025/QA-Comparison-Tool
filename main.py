import pandas as pd
import re
import os
from banner import print_banner  # Optional banner

# ─────────────────────────────────────────────
print_banner()

# File paths
file1 = "input/validated.xlsx"
file2 = "input/report.xlsx"

if not os.path.exists(file1) or not os.path.exists(file2):
    print("❌ One or both files are missing!")
    exit()

# ─────────────────────────────────────────────
# Load the validated workbook
df1_sheets = pd.ExcelFile(file1).sheet_names
requirements_sheet = next((s for s in df1_sheets if "requirements" in s.lower()), None)
if not requirements_sheet:
    print("❌ 'Requirements' sheet not found.")
    exit()

print(f"✅ Found sheet: {requirements_sheet}")
df1 = pd.read_excel(file1, sheet_name=requirements_sheet)
df1.columns = df1.columns.str.strip()

# ─────────────────────────────────────────────
# Helper: Find header row in report sheets
def find_header_row(df):
    for i, row in df.iterrows():
        if "Unique ID" in row.values:
            return i
    return 0

# ─────────────────────────────────────────────
# Load report workbook
df2_sheets = {}
for sheet_name, df in pd.read_excel(file2, sheet_name=None, header=None).items():
    header_row = find_header_row(df)
    df.columns = df.iloc[header_row].astype(str).str.strip()
    df = df[header_row + 1:].reset_index(drop=True)
    df2_sheets[sheet_name] = df

# ─────────────────────────────────────────────
# Extract score from report (inside parentheses)
def extract_report_score(value):
    if pd.isna(value):
        return value
    if isinstance(value, str):
        match = re.search(r"\((\d+)%\)", value)
        if match:
            return int(match.group(1))
    elif isinstance(value, (int, float)):
        return int(value)
    return None

# Extract plain score from validated
def extract_validated_score(value):
    if pd.isna(value):
        return value
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        match = re.search(r"\d+", value)
        if match:
            return int(match.group(0))
    return None

# ─────────────────────────────────────────────
# Define the specific column pairs to compare
column_pairs = [
    ("Maturity - Policy", "Adjusted Score - Policy", "Policy"),
    ("Maturity - Procedure", "Adjusted Score - Procedure", "Procedure"),
    ("Maturity - Implementation", "Adjusted Score - Implementation", "Implementation"),
    ("Maturity - Measured", "Adjusted Score - Measured", "Measured"),
    ("Maturity - Managed", "Adjusted Score - Managed", "Managed")
]

# ─────────────────────────────────────────────
# Normalize validated workbook 
df1["Unique ID"] = df1["Unique ID"].astype(str).str.strip().str.lower()
df1.columns = df1.columns.str.strip()

# Keep only needed columns from validated
needed_df1_cols = ["Unique ID"] + [v for _, v, _ in column_pairs]
df1 = df1[[col for col in needed_df1_cols if col in df1.columns]]

# ─────────────────────────────────────────────
# Begin comparison
all_differences = []

for sheet_name, df2 in df2_sheets.items():
    print(f"\n📄 Processing: {sheet_name}")
    
    df2.columns = df2.columns.str.strip()
    df2["Unique ID"] = df2["Unique ID"].astype(str).str.strip().str.lower()

    # Keep only needed columns from report
    needed_df2_cols = ["Unique ID"] + [r for r, _, _ in column_pairs]
    df2 = df2[[col for col in needed_df2_cols if col in df2.columns]]

    # Extract numbers in report columns (inside parentheses)
    for report_col, _, _ in column_pairs:
        if report_col in df2.columns:
            df2[report_col] = df2[report_col].apply(extract_report_score)

    # Extract numbers in validated columns (normal integers)
    for _, validated_col, _ in column_pairs:
        if validated_col in df1.columns:
            df1[validated_col] = df1[validated_col].apply(extract_validated_score)

    # Merge the data on Unique ID
    merged_df = pd.merge(df1, df2, on="Unique ID", how="inner", suffixes=('_validated', '_report'))
    print(f"🔗 Merged rows: {len(merged_df)}")

    # Compare each specified pair and log differences
    for _, row in merged_df.iterrows():
        for report_col, validated_col, display_name in column_pairs:
            val = row.get(validated_col)
            rep = row.get(report_col)

            if pd.notna(val) and pd.notna(rep) and val != rep:
                all_differences.append({
                    "Sheet Name": sheet_name,
                    "Unique ID": row["Unique ID"],
                    "Column": display_name,
                    "Report Value": rep,
                    "Validated Value": val
                })

# ─────────────────────────────────────────────
# Output results
diff_df = pd.DataFrame(all_differences)
output_file = "comparison_results.xlsx"
diff_df.to_excel(output_file, index=False)

if not diff_df.empty:
    print(f"\n🚨 Differences found! Saved to: {output_file}")
else:
    print("\n✅ No differences found!")

print("\n📊 Summary of Differences:")
print(diff_df)
