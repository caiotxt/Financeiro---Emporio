import os
import pandas as pd
import numpy as np

# ======================================================
# PUBLIC VERSION
# Mercado Livre SKU-level financial pipeline (generic)
# Business-sensitive logic was intentionally removed.
# ======================================================

BASE_MONTHS_DIR = "data/marketplace/months"
COSTS_PATH = "data/reference/costs.xlsx"
COSTS_SHEET = "PRODUCTS"
OUTPUT_DIR = "output"
DEFAULT_YEAR = 2025

PRODUCT_COL = "Operation description"

MONTH_MAP = {
    "01": ("JAN", 1), "02": ("FEV", 2), "03": ("MAR", 3), "04": ("ABR", 4),
    "05": ("MAI", 5), "06": ("JUN", 6), "07": ("JUL", 7), "08": ("AGO", 8),
    "09": ("SET", 9), "10": ("OUT", 10), "11": ("NOV", 11), "12": ("DEZ", 12),
}


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def clean_sku(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    return s if s else np.nan


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def extract_product_by_sku(df, sku_col, product_col):
    aux = (
        df[[sku_col, product_col]]
        .dropna(subset=[sku_col, product_col])
        .astype(str)
    )

    return (
        aux.groupby(sku_col)[product_col]
        .agg(lambda x: x.mode().iloc[0] if not x.mode().empty else x.iloc[-1])
        .reset_index()
        .rename(columns={
            sku_col: "SKU",
            product_col: "product"
        })
    )


# ------------------------------------------------------------------
# Core processing (generic logic only)
# ------------------------------------------------------------------
def process_collection(collection_path):
    df = pd.read_excel(collection_path)

    df["SKU"] = df["Seller SKU"].apply(clean_sku)

    if PRODUCT_COL not in df.columns:
        raise ValueError(f"Product column not found: {PRODUCT_COL}")

    product_sku = extract_product_by_sku(
        df,
        sku_col="SKU",
        product_col=PRODUCT_COL
    )

    df["net_amount"] = pd.to_numeric(
        df["Net received amount"], errors="coerce"
    ).fillna(0)

    df["gross_amount"] = pd.to_numeric(
        df["Transaction amount"], errors="coerce"
    ).fillna(0)

    df["quantity"] = np.where(
        df["gross_amount"] > 0, 1,
        np.where(df["gross_amount"] < 0, -1, 0)
    )

    sku_qty = (
        df.dropna(subset=["SKU"])
          .groupby("SKU", as_index=False)["quantity"]
          .sum()
          .query("quantity != 0")
    )

    revenue_sku = (
        df.dropna(subset=["SKU"])
          .groupby("SKU", as_index=False)["net_amount"]
          .sum()
          .rename(columns={"net_amount": "net_revenue"})
    )

    gross_sku = (
        df.dropna(subset=["SKU"])
          .groupby("SKU", as_index=False)["gross_amount"]
          .sum()
          .rename(columns={"gross_amount": "gross_revenue"})
    )

    # --------------------------------------------------
    # Cost reference (abstracted)
    # --------------------------------------------------
    costs = pd.read_excel(COSTS_PATH, sheet_name=COSTS_SHEET)
    costs = costs.rename(columns={costs.columns[0]: "SKU"})
    costs["SKU"] = costs["SKU"].apply(clean_sku)

    cost_col = costs.columns[1]
    costs[cost_col] = pd.to_numeric(costs[cost_col], errors="coerce")

    cost_ref = (
        costs[["SKU", cost_col]]
        .dropna(subset=["SKU"])
        .drop_duplicates(subset=["SKU"], keep="last")
        .rename(columns={cost_col: "unit_cost_estimated"})
    )

    result = (
        sku_qty
        .merge(revenue_sku, on="SKU", how="left")
        .merge(gross_sku, on="SKU", how="left")
        .merge(cost_ref, on="SKU", how="left")
        .merge(product_sku, on="SKU", how="left")
    )

    result["estimated_cmv"] = result["quantity"] * result["unit_cost_estimated"]

    # Profit and margin intentionally removed
    result["profit"] = np.nan
    result["margin"] = np.nan

    return result


# ------------------------------------------------------------------
# Batch processing
# ------------------------------------------------------------------
ensure_dir(OUTPUT_DIR)

for month_folder in sorted(os.listdir(BASE_MONTHS_DIR)):
    if month_folder not in MONTH_MAP:
        continue

    month_name, month_num = MONTH_MAP[month_folder]
    full_path = os.path.join(BASE_MONTHS_DIR, month_folder)

    files = [f for f in os.listdir(full_path) if f.endswith(".xlsx")]
    if not files:
        print(f"No files found for {month_folder}")
        continue

    collection_path = os.path.join(full_path, files[0])
    print(f"Processing {month_name}/{DEFAULT_YEAR}: {collection_path}")

    sku_result = process_collection(collection_path)

    output_name = f"ML_Report_{month_name}{str(DEFAULT_YEAR)[-2:]}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_name)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        sku_result.to_excel(writer, index=False, sheet_name="SKU_Financials")

    print(f"Generated: {output_name}")

print("\nAll months processed successfully.")
