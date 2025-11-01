#!/bin/bash

VERSION=3

# Define input file
INPUT_FILE="../Inventory_in.xlsx"

# Get list of sheet names as an array
mapfile -t SHEETS < <(
INPUT_FILE="$INPUT_FILE" python3 - <<'EOF'
import pandas as pd, os
xls = pd.ExcelFile(os.environ["INPUT_FILE"])
for sheet in xls.sheet_names:
    print(sheet)
EOF
"$INPUT_FILE"
)

# Loop over each sheet and call the reclassify script
echo "Start looping"
for i in  "${SHEETS[@]}"; do
  # Replace spaces in sheet names with underscores for safe filenames
    SAFE_NAME="${i// /}"
    OUT_FILE="../Inventory_out_${SAFE_NAME}_${VERSION}.xlsx"
    echo "Processing sheet: $i into file ${OUT_FILE} from ${INPUT_FILE}"
    python3 ./reclassify_tab_v5b.py --input "$INPUT_FILE" --sheet "$i" --output "$OUT_FILE"
    echo "Sheet $i completed"
done
