import streamlit as st
import pandas as pd
from analysis_v6 import run_analysis
from io import BytesIO
import plotly.express as px
from pathlib import Path
import zipfile
from zipfile import ZipFile
import tempfile

if "analysis_done" not in st.session_state:
    st.session_state.analysis_done = False


st.title("Experiment Analysis")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx"]
)

layout_text = st.text_area(
    "Paste pod loading scheme (tab-separated)\n"
    "Use the following format: concentration_condition",
    height=200,
    placeholder="T1\tT1\tT2\tT2\nT3\tT3\tT4\tT4"
)


if uploaded_file and layout_text:
    if st.button("Run analysis"):
        df = pd.read_excel(uploaded_file)
        layout_lines = layout_text.strip().split("\n")

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

    ## NEW
    # Add numeric helper column for sorting
    for df in [flat_ch2, flat_ch3]:
        df["Loaded_num"] = df["Loaded"].str.extract(r"(\d+)").astype(float)

    # Determine descending order per dataframe
    ch2_order = flat_ch2.sort_values(["Loaded_num", "Loaded"], ascending=[False, True])["Loaded"].unique()
    ch3_order = flat_ch3.sort_values(["Loaded_num", "Loaded"], ascending=[False, True])["Loaded"].unique()
    ##

    figures = {}

    # --- CH2 boxplots ---
    for metric in ["Cq", "Ampl.", "Slope"]:
        fig = px.box(
            ch2,
            x="Loaded",
            y=metric,
            points="all",
            title=f"{metric} by Loaded (CH2)",
            category_orders={"Loaded": ch2_order}
        )
        figures[f"CH2_{metric}_box"] = fig

    # --- CH3 boxplots ---
    for metric in ["Cq", "Ampl.", "Slope"]:
        fig = px.box(
            ch3,
            x="Loaded",
            y=metric,
            points="all",
            title=f"{metric} by Loaded (CH3)",
            category_orders={"Loaded": ch3_order}
        )
        figures[f"CH3_{metric}_box"] = fig

    # --- Detection rate ---
    figures["CH2_detection"] = px.bar(
        flat_ch2,
        x="Loaded",
        y="Detection_%_",
        range_y=[0, 1],
        text="Detection_%_",
        title="Detection rate (CH2)",
        category_orders={"Loaded": ch2_order}
    )
    # Percentage inside the bar
    figures["CH2_detection"].update_traces(
        text=flat_ch2["Detection_%_"].round(1).astype(str) + "%",
        textposition="inside"
    )
    # n on top of the bar
    for i, row in flat_ch2.iterrows():
        figures["CH2_detection"].add_annotation(
            x=row["Loaded"],
            y=row["Detection_%_"],
            text=f"n={int(row['QC_N_loaded'])}",
            showarrow=False,
            yshift=12
        )
    figures["CH2_detection"].update_yaxes(range=[0, 100])

    figures["CH3_detection"] = px.bar(
        flat_ch3,
        x="Loaded",
        y="Detection_%_",
        range_y=[0, 1],
        text="Detection_%_",
        title="Detection rate (CH3)",
        category_orders={"Loaded": ch3_order}
    )
    # Percentage inside the bar
    figures["CH3_detection"].update_traces(
        text=flat_ch3["Detection_%_"].round(1).astype(str) + "%",
        textposition="inside"
    )
    # n on top of the bar
    for i, row in flat_ch3.iterrows():
        figures["CH3_detection"].add_annotation(
            x=row["Loaded"],
            y=row["Detection_%_"],
            text=f"n={int(row['QC_N_loaded'])}",
            showarrow=False,
            yshift=12
        )
    figures["CH3_detection"].update_yaxes(range=[0, 100])


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