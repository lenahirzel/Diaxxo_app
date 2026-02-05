import pandas as pd
import numpy as np

def get_multiline_input(prompt=""):
    print(prompt)
    print("Paste your table below. When finished, type END on its own line.\n")

    lines = []
    while True:
        line = input()
        if line.strip() == "END":
            break
        lines.append(line)
    return "\n".join(lines)


def parse_pasted_qpcr(raw_text):

    if not raw_text.strip():
        print("❗ No data pasted. Aborting.")
        return None

    # Clean lines
    lines = [l.strip() for l in raw_text.splitlines() if l.strip()]

    # Check format
    if len(lines) % 9 != 0:
        print(f"⚠️ Warning: {len(lines)} lines found, not divisible by 9 (block size).")
        print("   The parser will proceed, but results may be misaligned.\n")

    # Expected 9-column block
    columns = [
        "Sample Name",
        "Well ID",
        "Channel",
        "Assay",
        "Cq",
        "Ampl.",
        "Slope",
        "Block",
        "Classification"
    ]

    blocks = [lines[i:i+9] for i in range(0, len(lines), 9)]
    df = pd.DataFrame(blocks, columns=columns)

    # Convert numeric
    for col in ["Cq", "Ampl.", "Slope"]:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # CH3 inherits sample name + well ID from preceding CH2
    for i in range(len(df)):
        if df.loc[i, "Channel"] == "CH3":
            df.loc[i, "Sample Name"] = df.loc[i-1, "Sample Name"]
            df.loc[i, "Well ID"] = df.loc[i-1, "Well ID"]

    # Detection calculations
    df["Detection_bool"] = df["Classification"].str.upper() == "POSITIVE"
    df["Detection (%)"] = df.groupby(["Sample Name", "Well ID"])["Detection_bool"].transform("mean") * 100

    # Split channels
    df_CH2 = df[df["Channel"] == "CH2"].copy()
    df_CH3 = df[df["Channel"] == "CH3"].copy()

    # Aggregations
    agg_funcs = {
        "Cq": ["mean", "std"],
        "Ampl.": ["mean", "std"],
        "Slope": ["mean", "std"],
        "Detection (%)": "mean"
    }

    stats_CH2 = df_CH2.groupby(["Sample Name", "Well ID"]).agg(agg_funcs)
    stats_CH3 = df_CH3.groupby(["Sample Name", "Well ID"]).agg(agg_funcs)

    # Flatten columns
    stats_CH2.columns = ['_'.join(col).replace(" ", "") for col in stats_CH2.columns]
    stats_CH3.columns = ['_'.join(col).replace(" ", "") for col in stats_CH3.columns]

    return df, df_CH2, df_CH3, stats_CH2, stats_CH3


# ----------------- MAIN -----------------

raw_text = get_multiline_input(
    "📋 Please paste the raw copied qPCR output table (from website)."
)

result = parse_pasted_qpcr(raw_text)

if result is not None:
    df, df_CH2, df_CH3, stats_CH2, stats_CH3 = result

    output_file = "qpcr_analysis.xlsx"
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Raw_Data", index=False)
        df_CH2.to_excel(writer, sheet_name="CH2_Data", index=False)
        df_CH3.to_excel(writer, sheet_name="CH3_Data", index=False)
        stats_CH2.to_excel(writer, sheet_name="CH2_Stats")
        stats_CH3.to_excel(writer, sheet_name="CH3_Stats")

    print("\n✅ Export complete → qpcr_analysis.xlsx")
    print("You can now open the Excel file.")