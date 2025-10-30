#!/usr/bin/env python3
"""
Reclassify a specific Excel sheet using Taxonomy v5b logic.

Usage:
    python reclassify_tab_v5b.py \
        --input "Inventory SY Smith Phase 1 for import v2_reclassified.xlsx" \
        --sheet "Salon" \
        --output "Inventory_SYSmith_v5b_Salon_Reclassified.xlsx" \
        --zip
"""

import argparse
import pandas as pd
import numpy as np
from pathlib import Path
import zipfile


# --------------------------------------------------------------------
# Core classification logic (Taxonomy v5b simplified for portability)
# --------------------------------------------------------------------

def text_fields(row, cols):
    vals = []
    for c in cols:
        if c in row and pd.notna(row[c]):
            vals.append(str(row[c]).lower())
    return " ".join(vals)


def assign_v5b(row):
    """Return (v5b_Category, v5b_SubCategory, v5b_Sub-SubCategory)"""
    orig_new_cat = str(row.get("New Category", "")).strip()
    orig_new_sub = str(row.get("New Sub-Category", "")).strip()
    orig_subcat = str(row.get("Subcategory", "")).strip()
    desc = text_fields(row, ["Item", "Description", "Subcategory"])

    vcat, vsub, vsubsub = orig_new_cat, orig_new_sub, str(row.get("New Sub-Sub-Category", "")).strip()

    # Simple keyword-based classification rules for portability
    if "clean" in desc or "polish" in desc:
        return "Common Maintenance", "Cleaning Equipment & Supplies", vsubsub
    if any(k in desc for k in ["book", "log", "paper", "record"]):
        return "Recreational", "Cabin & Office Supplies", vsubsub
    if any(k in desc for k in ["food", "snack", "galley", "kitchen"]):
        return "Recreational", "Galley Consumables", vsubsub
    if any(k in desc for k in ["sofa", "chair", "table", "lamp", "cushion", "pillow", "furniture"]):
        return "Recreational", "Recreational Components", vsubsub
    if any(k in desc for k in ["lubricant", "oil", "wax", "adhesive", "protectant"]):
        return "Recreational", "Recreational Consumables", vsubsub

    # Default: keep existing
    return vcat, vsub, vsubsub


def reclassify_tab_v5b(input_path, sheet_name, output_path):
    """Reclassify a single sheet and output new workbook with summaries."""
    df = pd.read_excel(input_path, sheet_name=sheet_name)

    # Ensure required columns exist
    for col in ["New Category", "New Sub-Category", "New Sub-Sub-Category"]:
        if col not in df.columns:
            df[col] = ""

    # Apply mapping
    v5b_results = df.apply(assign_v5b, axis=1, result_type="expand")
    v5b_results.columns = ["v5b_Category", "v5b_SubCategory", "v5b_Sub-SubCategory"]

    # Insert columns to the right of "New Sub-Sub-Category"
    cols = list(df.columns)
    insert_at = cols.index("New Sub-Sub-Category") + 1 if "New Sub-Sub-Category" in cols else len(cols)
    left, right = cols[:insert_at], cols[insert_at:]
    df_v5b = pd.concat([df[left], v5b_results, df[right]], axis=1)

    # Summaries
    total_rows = len(df_v5b)
    changed = (
        (df_v5b["v5b_Category"].fillna("") != df_v5b["New Category"].fillna("")) |
        (df_v5b["v5b_SubCategory"].fillna("") != df_v5b["New Sub-Category"].fillna(""))
    )
    changed_count = int(changed.sum())
    blank_count = int(df_v5b["v5b_SubCategory"].isna().sum())
    unique_v5b = df_v5b["v5b_Category"].nunique()

    summary1 = pd.DataFrame([{
        "Total Rows": total_rows,
        "Changed Rows": changed_count,
        "Changed %": round(100 * changed_count / max(total_rows, 1), 2),
        "Blank v5b Sub-Categories": blank_count,
        "v5b Categories Used": unique_v5b
    }])

    summary2 = (
        df_v5b[df_v5b["New Sub-Category"] == "Other Spare Parts"]
        .groupby(["New Category", "v5b_Category"])
        .size()
        .reset_index(name="Count")
    )

    # Write output workbook
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df_v5b.to_excel(writer, sheet_name=sheet_name, index=False)
        summary1.to_excel(writer, sheet_name="v5b_Summary1", index=False)
        summary2.to_excel(writer, sheet_name="v5b_Summary2", index=False)

    return output_path


# --------------------------------------------------------------------
# CLI Interface
# --------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Reclassify an Excel sheet using Taxonomy v5b logic."
    )
    parser.add_argument("--input", required=True, help="Path to the input Excel workbook.")
    parser.add_argument("--sheet", required=True, help="Name of the sheet to process.")
    parser.add_argument("--output", required=True, help="Path for the output Excel workbook.")
    parser.add_argument("--zip", action="store_true", help="Also create a .zip archive of the output file.")

    args = parser.parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    print(f"Processing sheet '{args.sheet}' from {input_path.name} ...")
    reclassify_tab_v5b(input_path, args.sheet, output_path)
    print(f"Reclassified workbook written to: {output_path}")

    if args.zip:
        zip_path = output_path.with_suffix(".zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(output_path, arcname=output_path.name)
        print(f"Compressed ZIP archive created: {zip_path}")

    print("✅ Done.")


if __name__ == "__main__":
    main()
