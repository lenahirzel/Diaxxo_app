import pandas as pd
import numpy as np

def add_replicate_count(summary, df_channel):
    # Count ALL rows per Loaded (independent of Cq or detection)
    n_loaded = (
        df_channel.groupby("Loaded")
        .size()
        .rename("N_loaded")
    )

    summary[("QC", "N_loaded")] = n_loaded
    return summary


def flatten_summary(summary, channel_name):
    flat = summary.copy()
    flat.columns = [
        "_".join(col) if isinstance(col, tuple) else col
        for col in flat.columns
    ]
    flat = flat.reset_index()
    flat["Channel"] = channel_name
    return flat

def run_analysis(df, layout_lines):
    # Clean columns
    df.columns = df.columns.str.strip().str.replace(" ", "_")

    df["Sample_Name"] = df["Sample_Name"].ffill()
    df["Well_ID"] = df["Well_ID"].ffill()

    # Parse layout
    layout = [row.split("\t") for row in layout_lines]
    num_rows = len(layout)
    num_cols = len(layout[0])

    mapping = {}
    well_number = 1
    for r in range(num_rows):
        for c in range(num_cols):
            mapping[well_number] = layout[r][c]
            well_number += 1

    df["Loaded"] = df["Well_ID"].map(mapping)

    # Split channels
    df_ch2 = df[df["Channel"] == "CH2"].copy()
    df_ch3 = df[df["Channel"] == "CH3"].copy()

    # Replace Cq = -1
    df_ch2.loc[df_ch2["Cq"] == -1, "Cq"] = np.nan
    df_ch3.loc[df_ch3["Cq"] == -1, "Cq"] = np.nan

    numeric_cols = ["Cq", "Ampl.", "Slope"]

    green_col = "Block02_Phase06_Cycle00_GREEN"
    if green_col in df.columns:
        df[green_col] = pd.to_numeric(df[green_col], errors="coerce")
        if df[green_col].notna().any():
            numeric_cols.append(green_col)

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

    summary_ch2 = add_replicate_count(summary_ch2, df_ch2)
    summary_ch3 = add_replicate_count(summary_ch3, df_ch3)

    # Detection %
    det_ch3 = (
        df_ch3.groupby("Loaded")["Classification"]
        .apply(lambda x: (x == "POSITIVE").mean() * 100)
    )
    summary_ch3["Detection_%"] = det_ch3

    det_ch2 = (
        df_ch2.groupby("Loaded")["Classification"]
        .apply(lambda x: (x == "POSITIVE").mean() * 100)
    )
    summary_ch2["Detection_%"] = det_ch2

    flat_ch2 = flatten_summary(summary_ch2, "CH2")
    flat_ch3 = flatten_summary(summary_ch3, "CH3")

    return df, summary_ch2, summary_ch3, flat_ch2, flat_ch3



