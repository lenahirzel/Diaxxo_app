import pandas as pd
import numpy as np
from openpyxl import load_workbook

# ---------------------------------------------------------
# 1. Load Excel file
# ---------------------------------------------------------
file_path = input("Enter path to your Excel file: ")
df = pd.read_excel(file_path)

df.columns = df.columns.str.strip().str.replace(" ", "_")

# Fill down Sample Name & Well_ID for the CH3 multiplex rows
df["Sample_Name"] = df["Sample_Name"].ffill()
df["Well_ID"] = df["Well_ID"].ffill()

print("\nColumns detected:", df.columns.tolist())


# ---------------------------------------------------------
# 2. Ask user to paste plate layout matrix
# ---------------------------------------------------------
print("\nPaste your plate layout (rows separated by newlines, columns by tabs):")
print("Example:\nT1\tT1\tT2\tT2\tT5\nT4\tT4\tT3\tT3\tT6\n...")

print("Paste now, and press ENTER twice when done.\n")

layout_lines = []
while True:
    line = input()
    if line.strip() == "":
        break
    layout_lines.append(line)

# Parse the layout into a matrix
layout = [row.split("\t") for row in layout_lines]

# Number of rows/columns
num_rows = len(layout)
num_cols = len(layout[0])

print(f"\nDetected layout with {num_rows} rows × {num_cols} columns")

# ---------------------------------------------------------
# 3. Map layout to well numbers (1–96 style)
# ---------------------------------------------------------
mapping = {}

well_number = 1
for r in range(num_rows):
    for c in range(num_cols):
        mapping[well_number] = layout[r][c]
        well_number += 1

print("\nCreated Well_ID → Loaded mapping:")
print(mapping)

df["Loaded"] = df["Well_ID"].map(mapping)

print("\n--- Added 'Loaded' column using layout ---")
print(df.head())


# ---------------------------------------------------------
# 4. Split CH2 vs CH3
# ---------------------------------------------------------
df_ch2 = df[df["Channel"] == "CH2"].copy()
df_ch3 = df[df["Channel"] == "CH3"].copy()

# ---------------------------------------------------------
# 5. Replace Cq = -1 with NaN (so they are excluded from mean/std)
# ---------------------------------------------------------
df_ch2.loc[df_ch2["Cq"] == -1, "Cq"] = np.nan
df_ch3.loc[df_ch3["Cq"] == -1, "Cq"] = np.nan

print(f"\nCH2: Replaced {(df[df['Channel'] == 'CH2']['Cq'] == -1).sum()} Cq values of -1 with NaN")
print(f"CH3: Replaced {(df[df['Channel'] == 'CH3']['Cq'] == -1).sum()} Cq values of -1 with NaN")

numeric_cols = ["Cq", "Ampl.", "Slope"]


# ---------------------------------------------------------
# 6. Aggregate CH2 and CH3 separately
# ---------------------------------------------------------
summary_ch2 = (
    df_ch2.groupby("Loaded")[numeric_cols]
    .agg(["mean", "std"])
    .sort_index()
)

summary_ch3 = (
    df_ch3.groupby("Loaded")[numeric_cols]
    .agg(["mean", "std"])
    .sort_index()
)

# ---------------------------------------------------------
# 7. Detection percentage (classification positive)
# ---------------------------------------------------------
detection = (
    df_ch3.groupby("Loaded")["Classification"]
    .apply(lambda x: (x == "POSITIVE").mean() * 100)
    .rename("Detection_%")
)
summary_ch3 = summary_ch3.assign(Detection_=detection)

detection = (
    df_ch2.groupby("Loaded")["Classification"]
    .apply(lambda x: (x == "POSITIVE").mean() * 100)
    .rename("Detection_%")
)
summary_ch2 = summary_ch2.assign(Detection_=detection)

print("\n--- CH2 Summary ---")
print(summary_ch2)

print("\n--- CH3 Summary ---")
print(summary_ch3)


# ---------------------------------------------------------
# 8. Save to SAME Excel file (append new sheets)
# ---------------------------------------------------------
with pd.ExcelWriter(
    file_path,
    engine="openpyxl",
    mode="a",                      # Append mode
    if_sheet_exists="replace"      # Replace sheet if it already exists
) as writer:
    summary_ch2.to_excel(writer, sheet_name="CH2_summary")
    summary_ch3.to_excel(writer, sheet_name="CH3_summary")
    df.to_excel(writer, sheet_name="Full_Data_Processed", index=False)

print(f"\n✅ Added analysis sheets to: {file_path}")
print("   • CH2_summary")
print("   • CH3_summary")
print("   • Full_Data_Processed")
