# Note: This script was used locally to merge raw per-day Excel exports
# (not included in this repo) into combined_channels.xlsx. Included here
# for transparency into the preprocessing pipeline.

import pandas as pd
from pathlib import Path

# Folder containing the raw per-day Excel exports
folder = Path("path/to/your/data")

# Find all .xlsx files
# Find all .xlsx files except temporary Excel lock files
files = sorted(
    f for f in folder.glob("*.xlsx")
    if not f.name.startswith("~$")
)
# Output file
output_file = folder / "combined_channels.xlsx"

# Store merged data
combined_data = {}

# Process Channel_1 -> Channel_10
for i in range(1, 11):
    sheet_name = f"Channel_{i}"
    dfs = []

    # Read same sheet from every workbook
    for file in files:
        print(f"Reading {sheet_name} from {file.name}")

        df = pd.read_excel(file, sheet_name=sheet_name)

        # Standardize column names
        df.columns = [
            "Int",
            "Ch",
            "Date",
            "Time",
            "Wheel (counts)",
            "Wheel Accum (counts)"
        ]

        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%m/%d/%Y")

        dfs.append(df)

    # Combine all data
    combined_df = pd.concat(dfs, ignore_index=True)

    # Create datetime for sorting
    combined_df["Datetime"] = pd.to_datetime(
        combined_df["Date"].astype(str) + " " +
        combined_df["Time"].astype(str),
        errors="coerce"
    )

    # Sort chronologically
    combined_df = combined_df.sort_values("Datetime")

    # Remove helper column
    combined_df = combined_df.drop(columns="Datetime")

    # Save into dictionary
    combined_data[sheet_name] = combined_df

# Write final workbook
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for sheet_name, df in combined_data.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"\nDone. Combined workbook saved to:\n{output_file}")
