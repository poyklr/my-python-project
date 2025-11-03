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
# Taxonomy v5b: Keyword Groups
# --------------------------------------------------------------------

electrical_consumable_keys = {
    "terminal","terminals","connector","connectors","fuse","fuses","label","labels",
    "zip tie","zip-tie","heat shrink","heat-shrink","lug","ring terminal","butt splice",
    "butt-splice","spade","boot","tie mount","cable tie","cable-tie","wire ferrule","ferrule"
}

electrical_component_keys = {
    "relay","relays","breaker","breakers","controller","controllers","converter","converters",
    "power supply","power supplies","switch","switches","sensor","sensors","outlet","outlets",
    "bus bar","disconnect","enclosure","box","contact","contactor","panel","display","gauge"
}

hull_maintenance_keys = {
    "adhesive","sealant","cleaning","painting","sanding","fiberglass","polish","tape",
    "protective","wrap","repair","supplies","paint","varnish","epoxy"
}

hull_rigging_keys = {
    "rigging","fitting","mooring","anchor","anchoring","line","rope","deck","fender",
    "winch","cover","sail","sails","halyard","sheet","bosun","mast","spool","traveler",
    "block","cleat","clutch","shackle","turnbuckle","stay"
}

recreational_water_keys = {"water","diving","snorkel","swim","surf","paddle","kite","sup"}
recreational_fitness_keys = {"fitness","gym","exercise","yoga","workout","weight","band","dumbbell","resistance"}

tools_keys = {"tool","tools","puller","pliers","screwdriver","drill","socket","wrench","punch","generator","clamp"}
test_measure_keys = {"pressure gauge","pressure gauges","gauge","measurement","calibration","manifold"}
cleaning_keys = {"cleaning","laundry","detergent","soap","polish","wipe","wipes","brush","broom","mop"}
ppe_keys = {"respiratory","respirator","uniform","uniforms","gloves","coverall","coveralls","mask","masks"}
docs_keys = {"manual","manuals","document","documents","logbook","logbooks","book","books"}

safety_fire_keys = {"fire","firefighting","extinguisher","fire hose","firehose","suppression"}
safety_rescue_keys = {"life jacket","lifejacket","harness","tether","survival","epirb","plb","flare","flares","liferaft","life raft"}

rec_components_keys = {"hinge","slide","drawer","latch","hardware","bracket","mount","knob","fixture"}
rec_consumables_keys = {"lubricant","lubricants","oil","adhesive","polish","protectant","wax"}


# --------------------------------------------------------------------
# v5b Reclassification Function
# --------------------------------------------------------------------

def assign_v5b(row):
    """Full Taxonomy v5b logic"""
    orig_new_cat = str(row.get("New Category", "")).strip()
    orig_new_sub = str(row.get("New Sub-Category", "")).strip()
    desc = " ".join(str(row.get(c, "")).lower() for c in ["Item", "Description", "Subcategory"])
    vsubsub = str(row.get("New Sub-Sub-Category", "")).strip()

    # ELECTRICAL
    if orig_new_cat == "Electrical" and orig_new_sub == "Other Spare Parts":
        if any(k in desc for k in electrical_consumable_keys):
            return "Electrical", "Electrical Consumables", vsubsub
        if any(k in desc for k in electrical_component_keys):
            return "Electrical", "Electrical Components & Devices", vsubsub
        return "Electrical", "Electrical Components & Devices", vsubsub

    # HULL
    if orig_new_cat == "Hull" and orig_new_sub == "Other Spare Parts":
        if any(k in desc for k in hull_maintenance_keys):
            return "Hull", "Hull Maintenance & Repair", vsubsub
        if any(k in desc for k in hull_rigging_keys):
            return "Hull", "Rigging & Deck Equipment", vsubsub
        if any(k in desc for k in safety_fire_keys):
            return "Safety", "Fire & Emergency Equipment", vsubsub
        if any(k in desc for k in safety_rescue_keys):
            return "Safety", "Rescue & Survival Equipment", vsubsub
        return "Hull", "Hull Maintenance & Repair", vsubsub

    # COMMON MAINTENANCE
    if orig_new_cat == "Common Maintenance" and orig_new_sub == "Other Spare Parts":
        if any(k in desc for k in tools_keys):
            return "Common Maintenance", "Tools & Equipment", vsubsub
        if any(k in desc for k in test_measure_keys):
            return "Common Maintenance", "Test & Measurement Tools", vsubsub
        if any(k in desc for k in cleaning_keys):
            return "Common Maintenance", "Cleaning Equipment & Supplies", vsubsub
        if any(k in desc for k in ppe_keys):
            return "Common Maintenance", "PPE & Uniforms", vsubsub
        if any(k in desc for k in docs_keys):
            return "Common Maintenance", "Documentation & Manuals", vsubsub
        return "Common Maintenance", "Maintenance Consumables", vsubsub

    # RECREATIONAL
    if orig_new_cat == "Recreational" and orig_new_sub == "Other Spare Parts":
        if any(k in desc for k in cleaning_keys):
            return "Common Maintenance", "Cleaning Equipment & Supplies", vsubsub
        if "tools" in desc:
            return "Common Maintenance", "Tools & Equipment", vsubsub
        if "medical" in desc or "safety" in desc:
            return "Safety", "First Aid & Emergency Equipment", vsubsub
        if "food" in desc or "galley" in desc:
            return "Recreational", "Galley Consumables", vsubsub
        if any(x in desc for x in ["office", "book", "document", "decor", "blanket", "logbook"]):
            return "Recreational", "Cabin & Office Supplies", vsubsub
        if any(k in desc for k in recreational_water_keys):
            return "Recreational", "Recreational Equipment", vsubsub
        if any(k in desc for k in recreational_fitness_keys):
            return "Recreational", "Sports & Fitness Equipment", vsubsub
        if any(k in desc for k in rec_components_keys):
            return "Recreational", "Recreational Components", vsubsub
        if any(k in desc for k in rec_consumables_keys):
            return "Recreational", "Recreational Consumables", vsubsub
        return "Recreational", "Recreational Cleaning & Storage", vsubsub

    # SAILING
    if orig_new_cat == "Sailing" and orig_new_sub == "Other Spare Parts":
        if any(x in desc for x in ["sail", "canvas", "cover", "stack pack", "stackpack"]):
            return "Sailing", "Sails & Canvas", vsubsub
        if any(x in desc for x in ["halyard", "sheet", "reef", "downhaul", "line", "rope"]):
            return "Sailing", "Running Rigging", vsubsub
        if any(x in desc for x in ["shroud", "stay", "turnbuckle", "chainplate", "standing"]):
            return "Sailing", "Standing Rigging", vsubsub
        if any(x in desc for x in ["winch", "traveler", "block", "cleat", "clutch"]):
            return "Sailing", "Winches & Deck Gear", vsubsub
        if any(x in desc for x in ["tape", "twine", "whipping", "grease", "lubricant"]):
            return "Sailing", "Sailing Consumables", vsubsub
        if any(x in desc for x in ["sewing", "palm", "needle", "machine"]):
            return "Sailing", "Sail Maintenance & Repair", vsubsub
        return "Sailing", "Sailing Consumables", vsubsub

    # SAFETY
    if orig_new_cat == "Safety" and orig_new_sub == "Other Spare Parts":
        if any(k in desc for k in safety_fire_keys):
            return "Safety", "Fire & Emergency Equipment", vsubsub
        if any(k in desc for k in safety_rescue_keys):
            return "Safety", "Rescue & Survival Equipment", vsubsub
        return "Safety", "First Aid & Emergency Equipment", vsubsub

    # DEFAULT: return unchanged
    return orig_new_cat, orig_new_sub, vsubsub


# --------------------------------------------------------------------
# Core function to process one sheet
# --------------------------------------------------------------------

def reclassify_tab_v5b(input_path: Path, sheet_name: str, output_path: Path):
    """Reclassify one sheet and write output workbook with summaries."""
    df = pd.read_excel(input_path, sheet_name=sheet_name)

    for col in ["New Category", "New Sub-Category", "New Sub-Sub-Category"]:
        if col not in df.columns:
            df[col] = ""

    results = df.apply(assign_v5b, axis=1, result_type="expand")
    results.columns = ["v5b_Category", "v5b_SubCategory", "v5b_Sub-SubCategory"]

    insert_at = df.columns.get_loc("New Sub-Sub-Category") + 1 if "New Sub-Sub-Category" in df.columns else len(df.columns)
    cols_left = list(df.columns[:insert_at])
    cols_right = list(df.columns[insert_at:])
    df_v5b = pd.concat([df[cols_left], results, df[cols_right]], axis=1)

    # Summaries
    total_rows = len(df_v5b)
    changed = (
        (df_v5b["v5b_Category"].fillna("") != df_v5b["New Category"].fillna("")) |
        (df_v5b["v5b_SubCategory"].fillna("") != df_v5b["New Sub-Category"].fillna(""))
    )
    summary1 = pd.DataFrame([{
        "Total Rows": total_rows,
        "Changed Rows": int(changed.sum()),
        "Changed %": round(100 * changed.sum() / max(total_rows, 1), 2),
        "Blank v5b Sub-Categories": int(df_v5b["v5b_SubCategory"].isna().sum()),
        "v5b Categories Used": df_v5b["v5b_Category"].nunique()
    }])

    summary2 = (
        df_v5b[df_v5b["New Sub-Category"] == "Other Spare Parts"]
        .groupby(["New Category", "v5b_Category"])
        .size()
        .reset_index(name="Count")
    )

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df_v5b.to_excel(writer, sheet_name=sheet_name, index=False)
        summary1.to_excel(writer, sheet_name=f"v5b_Summary1_{sheet_name}", index=False)
        summary2.to_excel(writer, sheet_name=f"v5b_Summary2_{sheet_name}", index=False)

    return output_path


# --------------------------------------------------------------------
# CLI Entry Point
# --------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Reclassify an Excel sheet using Taxonomy v5b.")
    parser.add_argument("--input", required=True, help="Input Excel file path")
    parser.add_argument("--sheet", required=True, help="Sheet name to process")
    parser.add_argument("--output", required=True, help="Output Excel file path")
    parser.add_argument("--zip", action="store_true", help="Create ZIP archive of the output file")

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
        print(f"ZIP archive created: {zip_path}")

    print("✅ Done.")


if __name__ == "__main__":
    main()
