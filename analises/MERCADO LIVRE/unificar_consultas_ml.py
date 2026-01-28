import os
import re
import pandas as pd

# ======================================================
# PUBLIC VERSION
# Mercado Livre monthly fact table (generic)
# Business-sensitive metrics were intentionally removed.
# ======================================================

DATA_DIR = "data/marketplace/output"
DEFAULT_YEAR = 2025

FACT_SHEET = "SKU_Financials"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "FACT_SKU_MONTHLY_ML.xlsx")

MONTH_MAP = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4,
    "MAI": 5, "JUN": 6, "JUL": 7, "AGO": 8,
    "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12
}


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def extract_period(filename):
    name = filename.upper()
    match = re.search(r"(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)", name)

    if not match:
        raise ValueError(f"Month not identified in filename: {filename}")

    month_abbr = match.group(1)
    month = MONTH_MAP[month_abbr]
    year = DEFAULT_YEAR
    period = f"{year}-{str(month).zfill(2)}"

    return year, month, period


def validate_columns(df, required_cols, context):
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in {context}: {missing}")


# ------------------------------------------------------------------
# Main processing
# ------------------------------------------------------------------
dfs = []

for file in os.listdir(DATA_DIR):
    if not file.endswith(".xlsx") or file.startswith("~$"):
        continue

    path = os.path.join(DATA_DIR, file)
    print(f"Processing: {file}")

    year, month, period = extract_period(file)

    df = pd.read_excel(path, sheet_name=FACT_SHEET)
    df.columns = [c.strip() for c in df.columns]

    required_columns = [
        "SKU",
        "quantity",
        "estimated_cmv"
    ]

    validate_columns(df, required_columns, f"{file} → {FACT_SHEET}")

    if "product" not in df.columns:
        df["product"] = None

    # BI normalization
    df = df.rename(columns={
        "SKU": "sku",
        "quantity": "quantity_sold",
        "estimated_cmv": "cmv_total"
    })

    # metadata
    df["year"] = year
    df["month"] = month
    df["period"] = period
    df["source_file"] = file

    # typing
    df["sku"] = df["sku"].astype(str)
    df["quantity_sold"] = df["quantity_sold"].astype(int)
    df["cmv_total"] = df["cmv_total"].astype(float)

    dfs.append(df)


fact_final = pd.concat(dfs, ignore_index=True)

# ordering
fact_final = fact_final.sort_values(["year", "month", "sku"])

# validation (SKU + period)
duplicates = fact_final.duplicated(subset=["sku", "period"])
if duplicates.any():
    raise ValueError("Duplicate SKU detected within the same period (Mercado Livre).")

fact_final.to_excel(OUTPUT_FILE, index=False)

print("\nProcess completed successfully.")
print(f"Output saved at:\n{OUTPUT_FILE}")
