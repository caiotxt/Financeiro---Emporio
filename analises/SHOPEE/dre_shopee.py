import argparse
import os
import re
from datetime import datetime
import pandas as pd
import numpy as np

# ======================================================
# PUBLIC VERSION
# Core business rules were intentionally abstracted.
# This script is for demonstration and portfolio purposes.
# ======================================================

DEFAULT_PARENT_SHEET = "Performance do Produto"
DEFAULT_SALES_SHEET  = "Visão Geral das Vendas"
DEFAULT_COSTS_SHEET  = "PRODUTOS"


# ------------------------------------------------------------------
# Business rules abstraction (real implementation is private)
# ------------------------------------------------------------------
def get_business_fees():
    """
    Placeholder for marketplace fees and fixed costs.
    Real values are part of the private service layer.
    """
    return {
        "variable_rate": 0.0,
        "fixed_fee_per_order": 0.0
    }


# ------------------------------------------------------------------
# Utility functions
# ------------------------------------------------------------------
def br_to_float(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return np.nan


def clean_sku(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    return s if s not in ["", "-", "nan", "None"] else np.nan


def safe_get(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    raise KeyError(f"Expected column not found: {candidates}")


def infer_period_label(df):
    if "Data" in df.columns and len(df) > 0:
        return str(df.loc[0, "Data"])
    return datetime.now().strftime("%Y-%m")


def normalize_period_for_filename(label):
    safe = re.sub(r"[^0-9A-Za-z_-]+", "_", label)
    return safe[:50]


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


# ------------------------------------------------------------------
# Core processing (generic logic only)
# ------------------------------------------------------------------
def process_month(parent_path, sales_path, costs_path, out_dir):
    parent = pd.read_excel(parent_path)
    sales  = pd.read_excel(sales_path)
    costs  = pd.read_excel(costs_path)

    sku_col = safe_get(parent, ["SKU", "SKU Principal", "SKU Principle"])
    qty_col = safe_get(parent, ["Unidades", "Units", "Units (Paid order)"])

    parent[sku_col] = parent[sku_col].apply(clean_sku)
    parent["units"] = pd.to_numeric(parent[qty_col], errors="coerce").fillna(0)

    sku_summary = (
        parent.groupby(sku_col, as_index=False)["units"]
        .sum()
        .query("units > 0")
    )

    cost_col = safe_get(costs, ["Custo", "COST"])
    costs[cost_col] = pd.to_numeric(costs[cost_col], errors="coerce")

    sku_summary = sku_summary.merge(
        costs[[sku_col, cost_col]],
        how="left",
        on=sku_col
    )

    sku_summary["estimated_cmv"] = sku_summary["units"] * sku_summary[cost_col]
    CMV_TOTAL = sku_summary["estimated_cmv"].sum()

    revenue_col = safe_get(sales, ["Vendas", "Sales"])
    orders_col  = safe_get(sales, ["Pedidos", "Orders"])

    REVENUE = br_to_float(sales.loc[0, revenue_col])
    ORDERS  = br_to_float(sales.loc[0, orders_col])

    fees = get_business_fees()

    VARIABLE_COSTS = REVENUE * fees["variable_rate"]
    FIXED_COSTS    = ORDERS * fees["fixed_fee_per_order"]

    # --------------------------------------------------
    # Profit calculation intentionally abstracted
    # --------------------------------------------------
    PROFIT = np.nan
    MARGIN = np.nan

    period_label = infer_period_label(sales)
    period_file  = normalize_period_for_filename(period_label)

    resumo = pd.DataFrame([{
        "Period": period_label,
        "Revenue": REVENUE,
        "CMV": CMV_TOTAL,
        "Variable Costs": VARIABLE_COSTS,
        "Fixed Costs": FIXED_COSTS,
        "Profit": PROFIT,
        "Margin": MARGIN
    }])

    ensure_dir(out_dir)
    out_path = os.path.join(out_dir, f"Financial_Report_{period_file}.xlsx")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        sku_summary.to_excel(writer, index=False, sheet_name="CMV_by_SKU")
        resumo.to_excel(writer, index=False, sheet_name="Summary")

    return out_path


# ------------------------------------------------------------------
# CLI
# ------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Generic financial analysis pipeline (public version)."
    )

    parser.add_argument("--parent", required=True, help="Parent SKU report")
    parser.add_argument("--sales", required=True, help="Sales overview report")
    parser.add_argument("--costs", required=True, help="Cost reference file")
    parser.add_argument("--outdir", default="output", help="Output directory")

    args = parser.parse_args()

    output = process_month(
        parent_path=args.parent,
        sales_path=args.sales,
        costs_path=args.costs,
        out_dir=args.outdir
    )

    print(f"Report generated: {output}")


if __name__ == "__main__":
    main()
