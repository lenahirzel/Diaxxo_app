import pandas as pd
import numpy as np


def add_replicate_count(summary, df_channel):
    """Add total number of loaded replicates per sample."""
    n_loaded = df_channel.groupby("Loaded").size().rename(("QC", "N_loaded"))
    summary[("QC", "N_loaded")] = n_loaded
    return summary


def flatten_summary(summary, channel_name):
    """Flatten multi-index summary table for export and plotting."""
    flat = summary.copy()

    flat.columns = [
        "_".join([str(x) for x in col if x != ""]) if isinstance(col, tuple) else str(col)
        for col in flat.columns
    ]

    flat = flat.reset_index()
    flat["Channel"] = channel_name

    # Extract metadata from Loaded name if formatted as 'conc_condition'
    if "Loaded" in flat.columns:
        parts = flat["Loaded"].astype(str).str.split("_", n=1, expand=True)
        flat["Concentration"] = parts[0]
        flat["Condition"] = parts[1] if parts.shape[1] > 1 else np.nan

    return flat


def run_analysis(df, layout_lines=None):
    """
    Analyze DiaxxoCare export file.

    Parameters
    ----------
    df : pandas.DataFrame
        Raw input dataframe.
    layout_lines : list[str], optional
        Plate layout mapping. If omitted, Sample_Name is used as Loaded.
    """

    # Clean column names
    df = df.copy()
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace(" ", "_", regex=False)
    )

    # Rename required CSV columns to internal app names
    rename_map = {
        "DPod_Well": "Well_ID",
        "Block02_Phase06_Cycle00_RGB_ch1": "Background",
        "DRFU": "Ampl",
        "RFIR": "Slope",
        "Cq": "Cq",
        "Channel": "Channel",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    required_columns = [
        "Well_ID",
        "Channel",
        "Background",
        "Ampl",
        "Slope",
        "Cq",
    ]

    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        raise ValueError(
            "Missing required CSV column(s): "
            + ", ".join(missing_columns)
            + ". Expected columns are: DPod Well, Channel, "
              "Block02_Phase06_Cycle00_RGB_ch1, DRFU, RFIR, Cq."
        )

    # Ensure numeric columns
    for col in ["Well_ID", "Cq", "Ampl", "Slope", "Background"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Define Loaded/sample identity
    if layout_lines is not None:
        layout = [row.split("\t") for row in layout_lines if row.strip()]
        mapping = {}
        well_number = 1

        for r in layout:
            for value in r:
                mapping[well_number] = value
                well_number += 1

        df["Loaded"] = df["Well_ID"].map(mapping)
    else:
        df["Loaded"] = df["Well_ID"].astype(str)

    # Metadata extraction
    loaded_parts = df["Loaded"].astype(str).str.split("_", n=1, expand=True)
    df["Concentration"] = loaded_parts[0]
    df["Condition"] = loaded_parts[1] if loaded_parts.shape[1] > 1 else np.nan

    # Split by channel
    df_ch2 = df[df["Channel"] == "CH2"].copy()
    df_ch3 = df[df["Channel"] == "CH3"].copy()

    # Invalid or non-detected Cq values
    for d in (df_ch2, df_ch3):
        if "Cq" in d.columns:
            d.loc[d["Cq"] <= 0, "Cq"] = np.nan

    ch2_raw = df_ch2.copy()
    ch3_raw = df_ch3.copy()

    # Metrics available for summarization
    numeric_cols = [c for c in ["Cq", "Ampl", "Slope", "Background"] if c in df.columns]

    # Summaries
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

    # Add replicate counts
    summary_ch2 = add_replicate_count(summary_ch2, df_ch2)
    summary_ch3 = add_replicate_count(summary_ch3, df_ch3)

    # Detection percentages based on Cq
    det_ch2 = (
        ch2_raw.groupby("Loaded")["Cq"]
        .apply(lambda x: (pd.to_numeric(x, errors="coerce") > 0).mean() * 100)
    )
    det_ch3 = (
        ch3_raw.groupby("Loaded")["Cq"]
        .apply(lambda x: (pd.to_numeric(x, errors="coerce") > 0).mean() * 100)
    )

    summary_ch2[("QC", "Detection_%")] = det_ch2
    summary_ch3[("QC", "Detection_%")] = det_ch3

    # Flatten for downstream plotting/export
    flat_ch2 = flatten_summary(summary_ch2, "CH2")
    flat_ch3 = flatten_summary(summary_ch3, "CH3")

    return df, ch2_raw, ch3_raw, flat_ch2, flat_ch3
