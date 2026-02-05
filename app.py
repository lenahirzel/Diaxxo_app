import streamlit as st
import pandas as pd
from analysis_v5 import run_analysis
from io import BytesIO
import plotly.express as px


if "analysis_done" not in st.session_state:
    st.session_state.analysis_done = False


st.title("Experiment Analysis")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx"]
)

layout_text = st.text_area(
    "Paste pod loading scheme (tab-separated)",
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
    st.dataframe(ch2)

    st.subheader("CH3 Summary")
    st.dataframe(ch3)

    st.header("Comparison by condition")

    channel = st.radio(
        "Channel",
        ["CH2", "CH3"],
        horizontal=True
    )

    metric = st.selectbox(
        "Select metric",
        ["Cq_mean", "Slope_mean", "Ampl._mean"]
    )

    plot_df = flat_ch2 if channel == "CH2" else flat_ch3

    fig = px.box(
        plot_df,
        x="Loaded",
        y=metric,
        points="all",
        title=f"{metric.replace('_', ' ')} by Loaded ({channel})"
    )

    st.plotly_chart(fig, use_container_width=True)

    st.header("Detection rate")

    # Calculate total replicates for annotation
    annotations_top = plot_df["QC_N_loaded"]

    fig_det = px.bar(
        plot_df,
        x="Loaded",
        y="Detection_%_",
        #text="Detection_%_",
        title=f"Detection % by Loaded ({channel})",
        text=plot_df["Detection_%_"]
    )

    # Show Detection % inside bar
    fig_det.update_traces(
        texttemplate="%{text:.1f}%",  # show as e.g., 85.3%
        textposition="inside",
        marker_color='steelblue'
    )

    # Add total n on top of bar
    for i, n in enumerate(annotations_top):
        fig_det.add_annotation(
            x=plot_df["Loaded"].iloc[i],
            y=plot_df["Detection_%_"].iloc[i] + 5,  # small offset above bar
            text=f"n={n}",
            showarrow=False,
            font=dict(size=12)
        )

    fig_det.update_yaxes(range=[0, 110])  # give room for n on top
    st.plotly_chart(fig_det, use_container_width=True)

    # Save to Excel for download
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        ch2.to_excel(writer, sheet_name="CH2_summary")
        ch3.to_excel(writer, sheet_name="CH3_summary")
        full_df.to_excel(writer, sheet_name="Full_Data_Processed", index=False)

    st.download_button(
        "Download Excel with analysis",
        data=output.getvalue(),
        file_name="qPCR_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
