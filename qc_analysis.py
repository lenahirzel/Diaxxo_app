import json
from pathlib import Path

import pandas as pd


QC_CONFIG_PATH = Path(__file__).with_name("qc_controls.json")

DEFAULT_QC_CONTROLS = [
    {
        "sample_name": "ASPC_5k",
        "expected_result": "positive",
        "cq_min": 28.0,
        "cq_max": 30.0,
        "channel": "CH3",
    },
    {
        "sample_name": "NTC",
        "expected_result": "negative",
        "cq_min": None,
        "cq_max": None,
        "channel": "CH3",
    },
]


def load_qc_controls():
    """Load saved QC controls from local JSON file, or create defaults."""
    if not QC_CONFIG_PATH.exists():
        save_qc_controls(DEFAULT_QC_CONTROLS)
        return DEFAULT_QC_CONTROLS

    with QC_CONFIG_PATH.open("r", encoding="utf-8") as file:
        return json.load(file)


def save_qc_controls(qc_controls):
    """Save QC controls to local JSON file."""
    with QC_CONFIG_PATH.open("w", encoding="utf-8") as file:
        json.dump(qc_controls, file, indent=2)


def qc_controls_to_dataframe(qc_controls):
    """Convert saved QC controls to a dataframe for Streamlit editing."""
    return pd.DataFrame(qc_controls)


def dataframe_to_qc_controls(qc_df):
    """Normalize edited QC controls dataframe back to serializable records."""
    qc_df = qc_df.copy()

    required_columns = [
        "sample_name",
        "expected_result",
        "cq_min",
        "cq_max",
        "channel",
    ]

    for column in required_columns:
        if column not in qc_df.columns:
            qc_df[column] = None

    qc_df = qc_df[required_columns]

    qc_df["sample_name"] = qc_df["sample_name"].astype(str).str.strip()
    qc_df["expected_result"] = (
        qc_df["expected_result"]
        .astype(str)
        .str.strip()
        .str.lower()
    )
    qc_df["channel"] = qc_df["channel"].astype(str).str.strip()

    qc_df["cq_min"] = pd.to_numeric(qc_df["cq_min"], errors="coerce")
    qc_df["cq_max"] = pd.to_numeric(qc_df["cq_max"], errors="coerce")

    qc_df = qc_df[qc_df["sample_name"] != ""]

    qc_df["cq_min"] = qc_df["cq_min"].where(pd.notna(qc_df["cq_min"]), None)
    qc_df["cq_max"] = qc_df["cq_max"].where(pd.notna(qc_df["cq_max"]), None)

    return qc_df.to_dict(orient="records")


def assess_qc_results(flat_ch2, flat_ch3, qc_controls):
    """
    Compare observed QC sample results with saved expectations.

    Positive expectation:
    - detection must be > 0
    - mean Cq must be inside cq_min/cq_max if provided

    Negative expectation:
    - detection must be 0
    - mean Cq should be empty/NaN
    """
    observed_by_channel = {
        "CH2": flat_ch2.copy(),
        "CH3": flat_ch3.copy(),
    }

    assessment_rows = []

    for control in qc_controls:
        sample_name = control["sample_name"]
        expected_result = str(control["expected_result"]).lower()
        channel = control.get("channel", "CH3")
        cq_min = control.get("cq_min")
        cq_max = control.get("cq_max")

        observed_df = observed_by_channel.get(channel, pd.DataFrame())

        if observed_df.empty or "Loaded" not in observed_df.columns:
            assessment_rows.append(
                {
                    "QC sample": sample_name,
                    "Channel": channel,
                    "Expected result": expected_result,
                    "Expected Cq range": format_cq_range(cq_min, cq_max),
                    "Observed detection %": None,
                    "Observed mean Cq": None,
                    "Result": "not found",
                    "QC assessment": "not passed",
                }
            )
            continue

        sample_rows = observed_df[observed_df["Loaded"] == sample_name]

        if sample_rows.empty:
            assessment_rows.append(
                {
                    "QC sample": sample_name,
                    "Channel": channel,
                    "Expected result": expected_result,
                    "Expected Cq range": format_cq_range(cq_min, cq_max),
                    "Observed detection %": None,
                    "Observed mean Cq": None,
                    "Result": "not found",
                    "QC assessment": "not passed",
                }
            )
            continue

        row = sample_rows.iloc[0]

        mean_cq = pd.to_numeric(row.get("Cq_mean"), errors="coerce")
        detection_percent = pd.to_numeric(row.get("QC_Detection_%"), errors="coerce")

        is_detected = pd.notna(detection_percent) and detection_percent > 0

        if expected_result == "positive":
            cq_in_range = True

            if cq_min is not None and pd.notna(cq_min):
                cq_in_range = cq_in_range and pd.notna(mean_cq) and mean_cq >= float(cq_min)

            if cq_max is not None and pd.notna(cq_max):
                cq_in_range = cq_in_range and pd.notna(mean_cq) and mean_cq <= float(cq_max)

            passed = is_detected and cq_in_range
            observed_result = "positive" if is_detected else "negative"

        elif expected_result == "negative":
            passed = not is_detected
            observed_result = "positive" if is_detected else "negative"

        else:
            passed = False
            observed_result = "invalid expected result"

        assessment_rows.append(
            {
                "QC sample": sample_name,
                "Channel": channel,
                "Expected result": expected_result,
                "Expected Cq range": format_cq_range(cq_min, cq_max),
                "Observed detection %": detection_percent,
                "Observed mean Cq": mean_cq,
                "Result": observed_result,
                "QC assessment": "passed" if passed else "not passed",
            }
        )

    return pd.DataFrame(assessment_rows)


def format_cq_range(cq_min, cq_max):
    if cq_min is None or pd.isna(cq_min):
        cq_min_text = ""
    else:
        cq_min_text = f"{float(cq_min):g}"

    if cq_max is None or pd.isna(cq_max):
        cq_max_text = ""
    else:
        cq_max_text = f"{float(cq_max):g}"

    if cq_min_text and cq_max_text:
        return f"{cq_min_text}-{cq_max_text}"

    if cq_min_text:
        return f">= {cq_min_text}"

    if cq_max_text:
        return f"<= {cq_max_text}"

    return "not applicable"