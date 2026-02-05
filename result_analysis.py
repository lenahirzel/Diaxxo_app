import pandas as pd

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
# Assuming your data uses numeric Well_IDs (1,2,3,...)
# Well numbering is row-major: 1 2 3 4 5 ... 12
mapping = {}

well_number = 1
for r in range(num_rows):
    for c in range(num_cols):
        mapping[well_number] = layout[r][c]
        well_number += 1

print("\nCreated Well_ID → Loaded mapping:")
print(mapping)

# Add to the dataframe
df["Loaded"] = df["Well_ID"].map(mapping)

print("\n--- Added 'Loaded' column using layout ---")
print(df.head())


# ---------------------------------------------------------
# 4. Split CH2 vs CH3
# ---------------------------------------------------------
df_ch2 = df[df["Channel"] == "CH2"]
df_ch3 = df[df["Channel"] == "CH3"]

numeric_cols = ["Cq", "Ampl.", "Slope"]


# ---------------------------------------------------------
# 5. Aggregate CH2 and CH3 separately
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
# 6. Detection percentage (classification positive)
# ---------------------------------------------------------
detection = (
    df_ch3.groupby("Loaded")["Classification"]
    .apply(lambda x: (x == "POSITIVE").mean() * 100)
    .rename("Detection_%")
)

# Add detection % as a column in the summary
summary_ch3 = summary_ch3.assign(Detection_=detection)

detection = (
    df_ch2.groupby("Loaded")["Classification"]
    .apply(lambda x: (x == "POSITIVE").mean() * 100)
    .rename("Detection_%")
)

# Add detection % as a column in the summary
summary_ch2 = summary_ch2.assign(Detection_=detection)

print("\n--- CH2 Summary ---")
print(summary_ch2)

print("\n--- CH3 Summary ---")
print(summary_ch3)


# ---------------------------------------------------------
# 6. Save Outputs
# ---------------------------------------------------------
output_file = "multiplex_qpcr_with_layout.xlsx"
with pd.ExcelWriter(output_file) as writer:
    summary_ch2.to_excel(writer, sheet_name="CH2_summary")
    summary_ch3.to_excel(writer, sheet_name="CH3_summary")
    df.to_excel(writer, sheet_name="Full_Data", index=False)

print(f"\nSaved results to: {output_file}")