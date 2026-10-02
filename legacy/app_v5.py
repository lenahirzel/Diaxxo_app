import streamlit as st
import pandas as pd
from analysis_v6 import run_analysis
from io import BytesIO, StringIO
import plotly.express as px
import plotly.graph_objects as go
from scipy.stats import ttest_ind
from pathlib import Path
import zipfile
from zipfile import ZipFile
import tempfile
import os

# --- Helpers ---

# def add_replicate_count(summary_df, df):
#     n_loaded = df.groupby("Loaded")["Cq"].count()
#     summary_df["QC_N_loaded"] = n_loaded
#     return summary_df
#
# def flatten_summary(summary_df, channel):
#     flat = summary_df.copy()
#     flat.columns = ["_".join(col).strip() if isinstance(col, tuple) else col for col in flat.columns.values]
#     flat["Channel"] = channel
#     return flat.reset_index()

def load_uploaded_csv_with_metadata(uploaded_file):
    raw_text = uploaded_file.getvalue().decode("utf-8-sig")
    lines = raw_text.splitlines()

    if len(lines) < 3:
        raise ValueError("CSV must contain at least 3 lines: machine ID, experiment ID, and header row.")

    machine_id = lines[0].strip()
    experiment_id = lines[1].strip()

    csv_data = "\n".join(lines[2:])
    df = pd.read_csv(StringIO(csv_data))
    df.columns = df.columns.str.strip().str.replace(" ", "_")

    df["Machine_ID"] = machine_id
    df["Experiment_ID"] = experiment_id

    return df, machine_id, experiment_id

# Force a colored template/palette for BOTH interactive display and static exports
px.defaults.template = "plotly_white"
px.defaults.color_discrete_sequence = px.colors.qualitative.Plotly


if "analysis_done" not in st.session_state:
    st.session_state.analysis_done = False

# --- Streamlit app ---

st.title("Experiment Analysis")

machine_id = st.text_input("Machine ID", placeholder="diaxxoCare123")
experiment_id = st.text_input("Experiment ID", placeholder="expID1322")

uploaded_file = st.file_uploader(
    "Upload CSV file",
    type=["csv"]
)

st.markdown(
    "Paste pod loading scheme (tab-separated)  \n"
    "Use the following format: `concentration_condition`"
)

layout_text = st.text_area(
    "pod_loading_scheme",
    height=200,
    placeholder="100_FluA\t100_FluA\t10_FluA\t10_FluA\n50_FluA\t50_FluA\t100_MG\t100_MG\t",
    label_visibility="collapsed",
)


if uploaded_file and layout_text:
    if st.button("Run analysis"):
        df, machine_id, experiment_id = load_uploaded_csv_with_metadata(uploaded_file)
        layout_lines = layout_text.strip().split("\n")

        # record metadata
        st.write(f"Machine ID: {machine_id}")
        st.write(f"Experiment ID: {experiment_id}")

        results = run_analysis(df, layout_lines)

        (
            st.session_state.full_df,
            st.session_state.ch2,
            st.session_state.ch3,
            st.session_state.flat_ch2,
            st.session_state.flat_ch3,
        ) = results

        st.session_state.analysis_done = True

if st.session_state.analysis_done:
    full_df = st.session_state.full_df
    ch2 = st.session_state.ch2
    ch3 = st.session_state.ch3
    flat_ch2 = st.session_state.flat_ch2
    flat_ch3 = st.session_state.flat_ch3

    st.success("Analysis completed!")

    st.subheader("CH2 Summary")
    st.dataframe(flat_ch2)

    st.subheader("CH3 Summary")
    st.dataframe(flat_ch3)

    # Add numeric helper column for sorting
    for df in [flat_ch2, flat_ch3]:
        df["Loaded_num"] = df["Loaded"].str.extract(r"(\d+)").astype(float)

    # Determine descending order per dataframe
    ch2_order = flat_ch2.sort_values(["Loaded_num", "Loaded"], ascending=[False, True])["Loaded"].unique()
    ch3_order = flat_ch3.sort_values(["Loaded_num", "Loaded"], ascending=[False, True])["Loaded"].unique()


    figures = {}

    # --- CH2 boxplots ---
    for metric in ["Cq", "Ampl.", "Slope"]:
        fig = px.box(
            ch2,
            x="Loaded",
            y=metric,
            color="Condition",
            points="all",
            title=f"{metric} by Loaded (CH2)",
            category_orders={"Loaded": ch2_order}
        )
        # cleaner look: legend not needed (x-axis already shows it)
        fig.update_layout(showlegend=False)
        figures[f"CH2_{metric}_box"] = fig

    # --- CH3 boxplots ---
    for metric in ["Cq", "Ampl.", "Slope"]:
        fig = px.box(
            ch3,
            x="Loaded",
            y=metric,
            color="Condition",
            points="all",
            title=f"{metric} by Loaded (CH3)",
            category_orders={"Loaded": ch3_order}
        )
        # cleaner look: legend not needed (x-axis already shows it)
        fig.update_layout(showlegend=False)
        figures[f"CH3_{metric}_box"] = fig

    # --- Detection rate ---
    det_col = "Detection_%_"

    # Make sure plotting columns are numeric (prevents weird label behavior)
    for df in (flat_ch2, flat_ch3):
        df[det_col] = pd.to_numeric(df.get(det_col), errors="coerce")
        df["QC_N_loaded"] = pd.to_numeric(df.get("QC_N_loaded"), errors="coerce")

    figures["CH2_detection"] = px.bar(
        flat_ch2,
        x="Loaded",
        y=det_col,
        range_y=[0, 100],
        color="Condition",
        title="Detection rate (CH2)",
        category_orders={"Loaded": ch2_order},
    )
    # % inside the bar (bigger font)
    figures["CH2_detection"].update_traces(
        texttemplate="%{y:.1f}%",
        textposition="inside",
        textfont_size=16,
        insidetextanchor="middle",
    )
    # n on top (separate text layer)
    figures["CH2_detection"].add_trace(
        go.Scatter(
            x=flat_ch2["Loaded"],
            y=flat_ch2[det_col],
            text=flat_ch2["QC_N_loaded"].apply(lambda v: f"n={int(v)}" if pd.notna(v) else ""),
            mode="text",
            textposition="top center",
            textfont=dict(size=14, color="black"),
            showlegend=False,
            cliponaxis=False,
        )
    )
    figures["CH2_detection"].update_layout(margin=dict(t=60))
    figures["CH2_detection"].update_yaxes(range=[0, 105])

    figures["CH3_detection"] = px.bar(
        flat_ch3,
        x="Loaded",
        y=det_col,
        range_y=[0, 100],
        color="Condition",
        title="Detection rate (CH3)",
        category_orders={"Loaded": ch3_order},
    )
    # % inside the bar (bigger font)
    figures["CH3_detection"].update_traces(
        texttemplate="%{y:.1f}%",
        textposition="inside",
        textfont_size=16,
        insidetextanchor="middle",
    )
    # n on top (separate text layer)
    figures["CH3_detection"].add_trace(
        go.Scatter(
            x=flat_ch3["Loaded"],
            y=flat_ch3[det_col],
            text=flat_ch3["QC_N_loaded"].apply(lambda v: f"n={int(v)}" if pd.notna(v) else ""),
            mode="text",
            textposition="top center",
            textfont=dict(size=14, color="black"),
            showlegend=False,
            cliponaxis=False,
        )
    )
    figures["CH3_detection"].update_layout(margin=dict(t=60))
    figures["CH3_detection"].update_yaxes(range=[0, 105])


    st.sidebar.header("Plots")

    show_ch2 = st.sidebar.checkbox("Show CH2", True)
    show_ch3 = st.sidebar.checkbox("Show CH3", True)
    show_detection = st.sidebar.checkbox("Show detection rate", True)

    if show_ch2:
        st.subheader("CH2")
        for key, fig in figures.items():
            if key.startswith("CH2_") and "detection" not in key:
                st.plotly_chart(fig, use_container_width=True)

    if show_ch3:
        st.subheader("CH3")
        for key, fig in figures.items():
            if key.startswith("CH3_") and "detection" not in key:
                st.plotly_chart(fig, use_container_width=True)

    if show_detection:
        st.subheader("Detection rate")
        if show_ch2:
            st.plotly_chart(figures["CH2_detection"], use_container_width=True)
        if show_ch3:
            st.plotly_chart(figures["CH3_detection"], use_container_width=True)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # ---- Save plots ----
        plot_dir = tmpdir / "plots"
        plot_dir.mkdir()

        for name, fig in figures.items():
            # Ensure exported images keep the same styling/colors
            fig.update_layout(template="plotly_white")
            fig.write_image(plot_dir / f"{name}.png", scale=2)

        # ---- Save Excel ----
        excel_path = tmpdir / "qPCR_analysis.xlsx"
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            flat_ch2.to_excel(writer, sheet_name="CH2_summary")
            flat_ch3.to_excel(writer, sheet_name="CH3_summary")
            full_df.to_excel(writer, sheet_name="Full_Data", index=False)

        # ---- Zip everything ----
        zip_buffer = BytesIO()
        with ZipFile(zip_buffer, "w") as zipf:
            zipf.write(excel_path, arcname="qPCR_analysis.xlsx")
            for img in plot_dir.iterdir():
                zipf.write(img, arcname=f"plots/{img.name}")

        st.download_button(
            "Download Excel + all plots",
            data=zip_buffer.getvalue(),
            file_name="qPCR_results.zip",
            mime="application/zip"
        )