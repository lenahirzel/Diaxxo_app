from io import BytesIO, StringIO
import csv
import streamlit as st
import pandas as pd
from analysis_v7 import run_analysis
from pod_to_pod_comparison_v2 import (
    run_pod_to_pod_comparison,
    figure_to_png_bytes,
    figure_to_pdf_bytes,
)
import plotly.express as px

def read_diaxxo_csv(uploaded_file):
    raw_text = uploaded_file.getvalue().decode("utf-8-sig", errors="replace")
    lines = raw_text.splitlines()

    header_row = None

    for idx, line in enumerate(lines):
        normalized_line = line.strip().replace(" ", "_")

        if "DPod_Well" in normalized_line:
            header_row = idx
            break

    if header_row is None:
        raise ValueError(
            "Could not find the real CSV header row. "
            "Expected a column named 'DPod Well'."
        )

    metadata_lines = lines[:header_row]
    csv_data = "\n".join(lines[header_row:])

    sample = csv_data[:4096]

    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,	,")
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = ";"

    df = pd.read_csv(
        StringIO(csv_data),
        sep=delimiter
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace(" ", "_", regex=False)
    )

    df = df.dropna(how="all")

    machine_id = metadata_lines[0].strip() if len(metadata_lines) > 0 else None
    experiment_id = metadata_lines[1].strip() if len(metadata_lines) > 1 else None

    df["Machine_ID"] = machine_id
    df["Experiment_ID"] = experiment_id

    return df


st.set_page_config(
    page_title="qPCR Pod Analysis",
    page_icon="🧪",
    layout="wide"
)


st.title("qPCR Pod Analysis App")

analysis_type = st.radio(
    "What would you like to do?",
    [
        "QC pod",
        "Comparison within one pod",
        "Comparison across multiple pods",
    ],
    index=None
)


if analysis_type is None:
    st.info("Please select an analysis option to continue.")
    st.stop()


uploaded_file = st.file_uploader(
    "Upload CSV file",
    type=["csv"]
)


if uploaded_file is None:
    st.info("Please upload a CSV file.")
    st.stop()


try:
    df = read_diaxxo_csv(uploaded_file)
except ValueError as error:
    st.error(str(error))
    st.stop()


if analysis_type == "QC pod":
    st.header("QC Pod")

    st.warning("QC pod analysis is not implemented yet.")

    st.subheader("Uploaded data preview")
    st.dataframe(df.head())


elif analysis_type == "Comparison within one pod":
    st.header("Comparison Within One Pod")

    st.markdown(
        "Paste pod loading scheme below.  \n"
        "Use the following format: `concentration_condition`"
    )

    layout_text = st.text_area(
        "Pod loading scheme",
        height=200,
        placeholder="100_FluA\t100_FluA\t10_FluA\t10_FluA\n50_FluA\t50_FluA\t100_MG\t100_MG"
    )

    if layout_text:
        if st.button("Run within-pod comparison"):
            layout_lines = layout_text.strip().split("\n")

            try:
                results = run_analysis(df, layout_lines)
            except ValueError as error:
                st.error(str(error))
                st.stop()

            (
                st.session_state.within_full_df,
                st.session_state.within_ch2,
                st.session_state.within_ch3,
                st.session_state.within_flat_ch2,
                st.session_state.within_flat_ch3,
            ) = results

            st.session_state.within_analysis_done = True

    else:
        st.info("Please paste the pod loading scheme.")

    if st.session_state.get("within_analysis_done", False):
        full_df = st.session_state.within_full_df
        ch2 = st.session_state.within_ch2
        ch3 = st.session_state.within_ch3
        flat_ch2 = st.session_state.within_flat_ch2
        flat_ch3 = st.session_state.within_flat_ch3

        st.success("Analysis completed!")

        st.subheader("CH2 Summary")
        st.dataframe(flat_ch2)

        st.subheader("CH3 Summary")
        st.dataframe(flat_ch3)

        for summary_df in [flat_ch2, flat_ch3]:
            summary_df["Loaded_num"] = (
                summary_df["Loaded"]
                .astype(str)
                .str.extract(r"(\d+)")
                .astype(float)
            )

        channel = st.radio(
            "Channel",
            ["CH2", "CH3"],
            horizontal=True,
            key="within_channel"
        )

        metric = st.selectbox(
            "Select metric",
            ["Cq_mean", "Slope_mean", "Ampl_mean", "Background_mean"],
            key="within_metric"
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

        fig_det = px.bar(
            plot_df,
            x="Loaded",
            y="QC_Detection_%",
            text="QC_Detection_%",
            title=f"Detection % by Loaded ({channel})"
        )

        fig_det.update_traces(
            texttemplate="%{text:.1f}%",
            textposition="inside",
            marker_color="steelblue"
        )

        fig_det.update_layout(
            yaxis_title="Detection %",
            xaxis_title="Loaded"
        )

        fig_det.update_yaxes(range=[0, 110])

        st.plotly_chart(fig_det, use_container_width=True)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            ch2.to_excel(writer, sheet_name="CH2_summary")
            ch3.to_excel(writer, sheet_name="CH3_summary")
            full_df.to_excel(writer, sheet_name="Full_Data_Processed", index=False)

        st.download_button(
            "Download Excel with analysis",
            data=output.getvalue(),
            file_name="within_pod_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


elif analysis_type == "Comparison across multiple pods":
    st.header("Comparison Across Multiple Pods")

    if st.button("Run across-pod comparison"):
        results = run_pod_to_pod_comparison(df)

        st.success("Across-pod comparison completed!")

        st.subheader("Detection Summary")
        st.dataframe(results["summary"])

        st.subheader("Processed Data")
        st.dataframe(results["df_all"])

        st.header("qPCR Figures")

        for channel, fig in results["publication_figures"].items():
            st.subheader(f"{channel} comparison")
            st.pyplot(fig)

            st.download_button(
                label=f"Download {channel} PNG",
                data=figure_to_png_bytes(fig),
                file_name=f"qpcr_{channel}.png",
                mime="image/png"
            )

            st.download_button(
                label=f"Download {channel} PDF",
                data=figure_to_pdf_bytes(fig),
                file_name=f"qpcr_{channel}.pdf",
                mime="application/pdf"
            )

        st.header("Detection Rate")

        detection_figure = results["detection_figure"]

        if detection_figure is not None:
            st.pyplot(detection_figure)

            st.download_button(
                label="Download detection rate PNG",
                data=figure_to_png_bytes(detection_figure),
                file_name="detection_rate.png",
                mime="image/png"
            )

            st.download_button(
                label="Download detection rate PDF",
                data=figure_to_pdf_bytes(detection_figure),
                file_name="detection_rate.pdf",
                mime="application/pdf"
            )
        else:
            st.warning("No CH3 data found for detection-rate plot.")