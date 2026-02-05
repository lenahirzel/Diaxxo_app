import streamlit as st
import pandas as pd
from analysis_v6 import run_analysis
from io import BytesIO
import plotly.express as px
from pathlib import Path
import zipfile

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

    st.header("Comparison by condition")

    channel = st.radio(
        "Channel",
        ["CH2", "CH3"],
        horizontal=True
    )

    # Use raw replicate-level dataframe
    plot_df = st.session_state.ch2 if channel == "CH2" else st.session_state.ch3

    metric_map = {
        "Cq": "Cq",
        "Amplitude": "Ampl.",
        "Slope": "Slope"
    }

    metric_label = st.selectbox(
        "Select metric",
        list(metric_map.keys())
    )

    metric_col = metric_map[metric_label]

    if channel == "CH2":
        fig_ch2 = px.box(
            ch2,
            x="Loaded",
            y=metric_col,
            points="all",
            title=f"{metric_label} by Loaded (CH2)"
        )
        st.plotly_chart(fig_ch2, use_container_width=True)

    elif channel == "CH3":
        fig_ch3 = px.box(
            ch3,
            x="Loaded",
            y=metric_col,
            points="all",
            title=f"{metric_label} by Loaded (CH3)"
        )
        st.plotly_chart(fig_ch3, use_container_width=True)


    st.header("Detection rate")

    det_df = flat_ch2 if channel == "CH2" else flat_ch3

    fig = px.bar(
        det_df,
        x="Loaded",
        y="Detection_%_",
        title=f"Detection % by Loaded ({channel})"
    )

    # Percentage inside the bar
    fig.update_traces(
        text=det_df["Detection_%_"].round(1).astype(str) + "%",
        textposition="inside"
    )

    # n on top of the bar
    for i, row in det_df.iterrows():
        fig.add_annotation(
            x=row["Loaded"],
            y=row["Detection_%_"],
            text=f"n={int(row['QC_N_loaded'])}",
            showarrow=False,
            yshift=12
        )

    fig.update_yaxes(range=[0, 100])

    st.plotly_chart(fig, use_container_width=True)

    # Save to Excel for download
    #output = BytesIO()
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        # Summary tables (mean / std / detection / n)
        flat_ch2.to_excel(writer, sheet_name="CH2_summary", index=False)
        flat_ch3.to_excel(writer, sheet_name="CH3_summary", index=False)

        # Raw replicate-level data
        #ch2.to_excel(writer, sheet_name="CH2_replicates", index=False)
       # ch3.to_excel(writer, sheet_name="CH3_replicates", index=False)

        # Full processed dataset
        full_df.to_excel(writer, sheet_name="Full_Data_Processed", index=False)

    excel_buffer.seek(0)

    plot_dir = Path("plots")
    plot_dir.mkdir(exist_ok=True)

    fig_ch2.write_image(plot_dir / "CH2_Cq_boxplot.png", scale=2)
    fig_ch3.write_image(plot_dir / "CH3_Cq_boxplot.png", scale=2)
    #det_ch2.write_image(plot_dir / "CH2_detection_rate.png", scale=2)
    #det_ch3.write_image(plot_dir / "CH3_detection_rate.png", scale=2)

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        # Add Excel
        zipf.writestr("summary.xlsx", excel_buffer.getvalue())

        # Add plots
        for img in plot_dir.glob("*.png"):
            zipf.write(img, arcname=f"plots/{img.name}")

    zip_buffer.seek(0)

    st.download_button(
        label="Download results (Excel + plots)",
        data=zip_buffer.getvalue(),
        file_name="results.zip",
        mime="application/zip"
    )
