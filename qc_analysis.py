import json
from pathlib import Path

import pandas as pd


QC_CONFIG_PATH = Path(__file__).with_name("qc_product_config.json")

def get_qc_config_path():
    """Return the absolute path of the QC configuration JSON file."""
    return QC_CONFIG_PATH.resolve()


DEFAULT_QC_CONFIG = {
    "244": {
        "assay_layout": [
            ["FluA", "FluA", "FluA", "FluA", "FluA"],
            ["H9", "H9", "H9", "H9", "H9"],
            ["b-actin", "b-actin", "b-actin", "b-actin", "b-actin"],
            ["FluA", "H9", "b-actin", "dxoPC", "dxoPC"],
        ],
        "qc_expectations": [
            {
                "qc_sample": "APS_5k",
                "assay": "FluA",
                "expected_result": "positive",
                "cq_min": 28.0,
                "cq_max": 30.0,
                "channel": "CH3",
            },
            {
                "qc_sample": "APS_5k",
                "assay": "H9",
                "expected_result": "positive",
                "cq_min": 30.0,
                "cq_max": 32.0,
                "channel": "CH3",
            },
            {
                "qc_sample": "NTC",
                "assay": "FluA",
                "expected_result": "negative",
                "cq_min": None,
                "cq_max": None,
                "channel": "CH3",
            },
            {
                "qc_sample": "NTC",
                "assay": "H9",
                "expected_result": "negative",
                "cq_min": None,
                "cq_max": None,
                "channel": "CH3",
            },
        ],
    }
}


def load_qc_config():
    """Load product-specific QC configuration from JSON, or create defaults."""
    if not QC_CONFIG_PATH.exists():
        save_qc_config(DEFAULT_QC_CONFIG)
        return DEFAULT_QC_CONFIG

    with QC_CONFIG_PATH.open("r", encoding="utf-8") as file:
        return json.load(file)


def save_qc_config(qc_config):
    """Save product-specific QC configuration to JSON."""
    with QC_CONFIG_PATH.open("w", encoding="utf-8") as file:
        json.dump(qc_config, file, indent=2)


def get_product_config(qc_config, product_number):
    """Return config for one product number, or an empty product config."""
    product_number = str(product_number).strip()

    return qc_config.get(
        product_number,
        {
            "assay_layout": [],
            "qc_expectations": [],
        },
    )


def assay_layout_to_text(assay_layout):
    """Convert nested assay layout list to tab-separated text."""
    if not assay_layout:
        return ""

    return "\n".join(
        "\t".join(str(value) for value in row)
        for row in assay_layout
    )


def assay_layout_text_to_rows(layout_text):
    """Convert tab-separated assay layout text to nested list."""
    return [
        [value.strip() for value in line.split("\t")]
        for line in layout_text.strip().split("\n")
        if line.strip()
    ]


def expectations_to_dataframe(qc_expectations):
    """Convert product-specific QC expectations to an editable dataframe."""
    columns = [
        "qc_sample",
        "assay",
        "expected_result",
        "cq_min",
        "cq_max",
        "channel",
    ]

    if not qc_expectations:
        return pd.DataFrame(columns=columns)

    return pd.DataFrame(qc_expectations).reindex(columns=columns)


def dataframe_to_expectations(expectations_df):
    """Normalize edited QC expectations dataframe back to serializable records."""
    expectations_df = expectations_df.copy()

    required_columns = [
        "qc_sample",
        "assay",
        "expected_result",
        "cq_min",
        "cq_max",
        "channel",
    ]

    for column in required_columns:
        if column not in expectations_df.columns:
            expectations_df[column] = None

    expectations_df = expectations_df[required_columns]

    expectations_df["qc_sample"] = expectations_df["qc_sample"].astype(str).str.strip()
    expectations_df["assay"] = expectations_df["assay"].astype(str).str.strip()
    expectations_df["expected_result"] = (
        expectations_df["expected_result"]
        .astype(str)
        .str.strip()
        .str.lower()
    )
    expectations_df["channel"] = expectations_df["channel"].astype(str).str.strip()

    expectations_df["cq_min"] = pd.to_numeric(expectations_df["cq_min"], errors="coerce")
    expectations_df["cq_max"] = pd.to_numeric(expectations_df["cq_max"], errors="coerce")

    expectations_df = expectations_df[
        (expectations_df["qc_sample"] != "")
        & (expectations_df["assay"] != "")
    ]

    expectations_df["cq_min"] = expectations_df["cq_min"].where(
        pd.notna(expectations_df["cq_min"]),
        None,
    )
    expectations_df["cq_max"] = expectations_df["cq_max"].where(
        pd.notna(expectations_df["cq_max"]),
        None,
    )

    return expectations_df.to_dict(orient="records")


def parse_sample_loading_scheme(sample_layout_text):
    """Convert pasted QC sample loading scheme to nested rows."""
    return [
        [value.strip() for value in line.split("\t")]
        for line in sample_layout_text.strip().split("\n")
        if line.strip()
    ]


def validate_matching_layout_shapes(sample_layout, assay_layout):
    """Ensure sample loading scheme and saved assay layout have the same shape."""
    if len(sample_layout) != len(assay_layout):
        raise ValueError(
            "The pasted sample loading scheme and saved assay layout have different row counts."
        )

    for row_index, sample_row in enumerate(sample_layout):
        assay_row = assay_layout[row_index]

        if len(sample_row) != len(assay_row):
            raise ValueError(
                "The pasted sample loading scheme and saved assay layout have different "
                f"column counts in row {row_index + 1}."
            )


def build_combined_layout_lines(sample_layout_text, assay_layout):
    """
    Combine pasted QC samples with saved assay layout.

    Example resulting Loaded value:
    APS_5k__FluA
    """
    sample_layout = parse_sample_loading_scheme(sample_layout_text)
    validate_matching_layout_shapes(sample_layout, assay_layout)

    combined_rows = []

    for sample_row, assay_row in zip(sample_layout, assay_layout):
        combined_row = []

        for qc_sample, assay in zip(sample_row, assay_row):
            combined_row.append(f"{qc_sample}__{assay}")

        combined_rows.append("\t".join(combined_row))

    return combined_rows


def add_qc_sample_and_assay_columns(*dataframes):
    """Add QC_sample and Assay columns from combined Loaded names."""
    for dataframe in dataframes:
        if dataframe is None or dataframe.empty or "Loaded" not in dataframe.columns:
            continue

        loaded_parts = dataframe["Loaded"].astype(str).str.split("__", n=1, expand=True)
        dataframe["QC_sample"] = loaded_parts[0]
        dataframe["Assay"] = loaded_parts[1] if loaded_parts.shape[1] > 1 else pd.NA


def assess_qc_results(flat_ch2, flat_ch3, qc_expectations):
    """
    Check each expected QC sample-assay combination against observed results.

    Positive expectation:
    - detection must be > 0
    - mean Cq must be inside cq_min/cq_max if provided

    Negative expectation:
    - detection must be 0
    """
    flat_ch2 = flat_ch2.copy()
    flat_ch3 = flat_ch3.copy()

    add_qc_sample_and_assay_columns(flat_ch2, flat_ch3)

    observed_by_channel = {
        "CH2": flat_ch2,
        "CH3": flat_ch3,
    }

    assessment_rows = []

    for expectation in qc_expectations:
        qc_sample = expectation["qc_sample"]
        assay = expectation["assay"]
        expected_result = str(expectation["expected_result"]).lower()
        channel = expectation.get("channel", "CH3")
        cq_min = expectation.get("cq_min")
        cq_max = expectation.get("cq_max")

        observed_df = observed_by_channel.get(channel, pd.DataFrame())

        if observed_df.empty:
            assessment_rows.append(
                build_missing_assessment_row(
                    qc_sample,
                    assay,
                    channel,
                    expected_result,
                    cq_min,
                    cq_max,
                )
            )
            continue

        sample_rows = observed_df[
            (observed_df["QC_sample"] == qc_sample)
            & (observed_df["Assay"] == assay)
        ]

        if sample_rows.empty:
            assessment_rows.append(
                build_missing_assessment_row(
                    qc_sample,
                    assay,
                    channel,
                    expected_result,
                    cq_min,
                    cq_max,
                )
            )
            continue

        row = sample_rows.iloc[0]

        mean_cq = pd.to_numeric(row.get("Cq_mean"), errors="coerce")
        detection_percent = pd.to_numeric(row.get("QC_Detection_%"), errors="coerce")
        n_loaded = pd.to_numeric(row.get("QC_N_loaded"), errors="coerce")

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
                "QC sample": qc_sample,
                "Assay": assay,
                "Channel": channel,
                "Expected result": expected_result,
                "Expected Cq range": format_cq_range(cq_min, cq_max),
                "Observed detection %": detection_percent,
                "Observed mean Cq": mean_cq,
                "N loaded": n_loaded,
                "Observed result": observed_result,
                "QC assessment": "passed" if passed else "not passed",
            }
        )

    return pd.DataFrame(assessment_rows)


def build_missing_assessment_row(qc_sample, assay, channel, expected_result, cq_min, cq_max):
    return {
        "QC sample": qc_sample,
        "Assay": assay,
        "Channel": channel,
        "Expected result": expected_result,
        "Expected Cq range": format_cq_range(cq_min, cq_max),
        "Observed detection %": None,
        "Observed mean Cq": None,
        "N loaded": None,
        "Observed result": "not found",
        "QC assessment": "not passed",
    }


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