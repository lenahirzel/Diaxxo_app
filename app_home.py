from io import BytesIO, StringIO
import csv
import streamlit as st
import pandas as pd
from analysis_v7 import run_analysis
from qc_analysis import (
    add_qc_sample_and_assay_columns,
    assay_layout_text_to_rows,
    assay_layout_to_text,
    assess_qc_results,
    build_combined_layout_lines,
    dataframe_to_expectations,
    expectations_to_dataframe,
    get_product_config,
    get_qc_config_path,
    load_qc_config,
    save_qc_config,
)
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


st.title("dPod Experiment Analysis App")

st.markdown(
    """
    **QC pod**  
    Check the quality of a single pod run. This option will be used for pod-level QC metrics and basic run validation.

    **Comparison within one pod**  
    Compare different loaded conditions within one pod. Enter loading scheme and CSV directly from single experiment.

    **Comparison across multiple pods**  
    Compare results between different pods or experiments. Use this when the CSV already contains data from multiple pods. Groups by condition. Required columns of the CSV: "Cq", "Ampl", "Slope", "Channel", "Condition", "Loaded"
    """
)

st.markdown("<br>", unsafe_allow_html=True)

st.markdown("### **What would you like to do?**")


analysis_descriptions = {
    "QC pod": (
        "Check the quality of a single pod run. "
        "This option will be used for pod-level QC metrics and basic run validation."
    ),
    "Comparison within one pod": (
        "Compare different loaded conditions within one pod. "
        "Use this when one CSV contains several concentrations or sample conditions."
    ),
    "Comparison across multiple pods": (
        "Compare results between different pods or experiments. "
        "Use this when the CSV already contains data from multiple pods."
    ),
}

analysis_type = st.radio(
    "Select analysis type",
    list(analysis_descriptions.keys()),
    index=None,
    label_visibility="collapsed"
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

    qc_config = load_qc_config()
    st.caption(f"QC config file: `{get_qc_config_path()}`")

st.subheader("LOT information")

    col1, col2 = st.columns(2)

    with col1:
        product_number = st.text_input("Product number")
        lot_sn = st.text_input("LOT serial number")

    with col2:
        manufacturing_date = st.date_input("Manufacturing date", value=None)
        expiration_date = st.date_input("Expiration date", value=None)

    product_number = str(product_number).strip()

    if not product_number:
        st.info("Please enter a product number to load the assay layout and QC expectations.")
        st.stop()

    product_config = get_product_config(qc_config, product_number)
    assay_layout = product_config.get("assay_layout", [])
    qc_expectations = product_config.get("qc_expectations", [])

    st.divider()

    st.subheader(f"Product {product_number} assay loading scheme")

    st.markdown(
        "This assay layout is linked to the product number. "
        "It defines which assay is present in each pod position."
    )

    assay_layout_text = st.text_area(
        "Assay loading scheme",
        value=assay_layout_to_text(assay_layout),
        height=160,
        placeholder="FluA\tFluA\tFluA\tFluA\tFluA\nH9\tH9\tH9\tH9\tH9\nb-actin\tb-actin\tb-actin\tb-actin\tb-actin\nFluA\tH9\tb-actin\tdxoPC\tdxoPC",
        key=f"assay_layout_{product_number}",
    )

    if st.button("Update assay loading scheme"):
        qc_config.setdefault(product_number, {})
        qc_config[product_number]["assay_layout"] = assay_layout_text_to_rows(assay_layout_text)
        qc_config[product_number]["qc_expectations"] = qc_expectations
        save_qc_config(qc_config)
        st.success(f"Assay loading scheme for product {product_number} was updated.")
        st.rerun()

    assay_layout = assay_layout_text_to_rows(assay_layout_text) if assay_layout_text.strip() else []

    st.divider()

    st.subheader(f"Product {product_number} QC expectations")

    st.markdown(
        "These expectations are specific to the combination of "
        "**product number + QC sample + assay + channel**."
    )

    expectations_df = expectations_to_dataframe(qc_expectations)

    edited_expectations_df = st.data_editor(
        expectations_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "qc_sample": st.column_config.TextColumn(
                "QC sample",
                help="Example: APS_5k, NTC",
                required=True,
            ),
            "assay": st.column_config.TextColumn(
                "Assay",
                help="Example: FluA, H9, b-actin",
                required=True,
            ),
            "expected_result": st.column_config.SelectboxColumn(
                "Expected result",
                options=["positive", "negative"],
                required=True,
            ),
            "cq_min": st.column_config.NumberColumn(
                "Minimum Cq",
                help="Leave empty if no lower limit is required.",
            ),
            "cq_max": st.column_config.NumberColumn(
                "Maximum Cq",
                help="Leave empty if no upper limit is required.",
            ),
            "channel": st.column_config.SelectboxColumn(
                "Channel",
                options=["CH2", "CH3"],
                required=True,
            ),
        },
        key=f"qc_expectations_{product_number}",
    )

    if st.button("Update QC expectations"):
        updated_expectations = dataframe_to_expectations(edited_expectations_df)
        qc_config.setdefault(product_number, {})
        qc_config[product_number]["assay_layout"] = assay_layout
        qc_config[product_number]["qc_expectations"] = updated_expectations
        save_qc_config(qc_config)
        st.success(f"QC expectations for product {product_number} were updated.")
        st.rerun()

    qc_expectations = dataframe_to_expectations(edited_expectations_df)

    saved_qc_samples = sorted(
        {
            expectation["qc_sample"]
            for expectation in qc_expectations
            if expectation.get("qc_sample")
        }
    )

    saved_assays = sorted(
        {
            expectation["assay"]
            for expectation in qc_expectations
            if expectation.get("assay")
        }
    )

    st.divider()

    st.subheader("QC sample loading scheme")

    st.markdown(
        "Paste the QC sample loading scheme below.  \n"
        "The app will combine this sample layout with the saved product assay layout."
    )

    if saved_qc_samples:
        st.info(
            "Saved QC samples for this product: "
            + ", ".join(f"`{sample}`" for sample in saved_qc_samples)
        )

    if saved_assays:
        st.info(
            "Assays with saved expectations for this product: "
            + ", ".join(f"`{assay}`" for assay in saved_assays)
        )

    sample_layout_text = st.text_area(
        "QC sample loading scheme",
        height=200,
        placeholder="APS_5k\tAPS_5k\tAPS_5k\tAPS_5k\tAPS_5k\nAPS_5k\tAPS_5k\tAPS_5k\tAPS_5k\tAPS_5k\nNTC\tNTC\tNTC\tNTC\tNTC\nAPS_5k\tAPS_5k\tAPS_5k\tNTC\tNTC",
        key=f"qc_sample_layout_{product_number}",
    )

    if sample_layout_text:
        loaded_qc_samples = [
            value.strip()
            for line in sample_layout_text.strip().split("\n")
            for value in line.split("\t")
            if value.strip()
        ]

        unknown_qc_samples = sorted(
            {
                value
                for value in loaded_qc_samples
                if value not in saved_qc_samples
            }
        )

        if unknown_qc_samples:
            st.warning(
                "The following loaded QC samples do not have saved expectations "
                f"for product {product_number}: "
                + ", ".join(f"`{value}`" for value in unknown_qc_samples)
            )

        if st.button("Run QC analysis"):
            if not assay_layout:
                st.error("No assay loading scheme is saved or entered for this product.")
                st.stop()

            try:
                combined_layout_lines = build_combined_layout_lines(
                    sample_layout_text,
                    assay_layout,
                )
            except ValueError as error:
                st.error(str(error))
                st.stop()

            try:
                results = run_analysis(df, combined_layout_lines)
            except ValueError as error:
                st.error(str(error))
                st.stop()

            (
                st.session_state.qc_full_df,
                st.session_state.qc_ch2,
                st.session_state.qc_ch3,
                st.session_state.qc_flat_ch2,
                st.session_state.qc_flat_ch3,
            ) = results

            add_qc_sample_and_assay_columns(
                st.session_state.qc_full_df,
                st.session_state.qc_ch2,
                st.session_state.qc_ch3,
                st.session_state.qc_flat_ch2,
                st.session_state.qc_flat_ch3,
            )

            st.session_state.qc_metadata = {
                "Product number": product_number,
                "LOT serial number": lot_sn,
                "Manufacturing date": manufacturing_date,
                "Expiration date": expiration_date,
            }

            st.session_state.qc_assay_layout = assay_layout
            st.session_state.qc_sample_layout_text = sample_layout_text
            st.session_state.qc_assessment = assess_qc_results(
                st.session_state.qc_flat_ch2,
                st.session_state.qc_flat_ch3,
                qc_expectations,
            )

            st.session_state.qc_analysis_done = True

    else:
        st.info("Please paste the QC sample loading scheme.")

    if st.session_state.get("qc_analysis_done", False):
        full_df = st.session_state.qc_full_df
        ch2 = st.session_state.qc_ch2
        ch3 = st.session_state.qc_ch3
        flat_ch2 = st.session_state.qc_flat_ch2
        flat_ch3 = st.session_state.qc_flat_ch3
        qc_assessment = st.session_state.qc_assessment
        qc_metadata = st.session_state.qc_metadata

        st.success("QC analysis completed!")

        overall_qc_passed = (
            not qc_assessment.empty
            and (qc_assessment["QC assessment"] == "passed").all()
        )

        if overall_qc_passed:
            st.success("Overall QC result: PASSED")
        else:
            st.error("Overall QC result: NOT PASSED")

        st.subheader("LOT information")

        metadata_df = pd.DataFrame(
            list(qc_metadata.items()),
            columns=["Field", "Value"],
        )

        st.dataframe(metadata_df, use_container_width=True)

        st.subheader("QC assessment by sample and assay")

        def highlight_qc_assessment(row):
            if row["QC assessment"] == "passed":
                return ["background-color: #d4edda"] * len(row)

            return ["background-color: #f8d7da"] * len(row)

        st.dataframe(
            qc_assessment.style.apply(highlight_qc_assessment, axis=1),
            use_container_width=True,
        )

        st.subheader("CH2 Summary")
        st.dataframe(flat_ch2, use_container_width=True)

        st.subheader("CH3 Summary")
        st.dataframe(flat_ch3, use_container_width=True)

        channel = st.radio(
            "Channel",
            ["CH2", "CH3"],
            horizontal=True,
            key="qc_channel",
        )

        box_plot_df = ch2 if channel == "CH2" else ch3
        detection_plot_df = flat_ch2 if channel == "CH2" else flat_ch3

        metric_options = {
            "Cq": "Cq",
            "Amplitude": "Ampl",
            "Slope": "Slope",
            "Background": "Background",
        }

        metric_label = st.selectbox(
            "Select metric",
            list(metric_options.keys()),
            key="qc_metric",
        )

        metric_column = metric_options[metric_label]

        if metric_column not in box_plot_df.columns:
            st.warning(
                f"No column found for {metric_label}. "
                f"Available columns: {', '.join(box_plot_df.columns)}"
            )
        else:
            box_plot_df[metric_column] = pd.to_numeric(
                box_plot_df[metric_column],
                errors="coerce",
            )

            fig = px.box(
                box_plot_df,
                x="Assay",
                y=metric_column,
                color="QC_sample",
                points="all",
                title=f"{metric_label} by assay and QC sample ({channel})",
            )

            fig.update_layout(
                xaxis_title="Assay",
                yaxis_title=metric_label,
            )

            st.plotly_chart(fig, use_container_width=True)

        st.header("Detection rate")

        fig_det = px.bar(
            detection_plot_df,
            x="Assay",
            y="QC_Detection_%",
            color="QC_sample",
            barmode="group",
            text="QC_Detection_%",
            title=f"Detection % by assay and QC sample ({channel})",
        )

        fig_det.update_traces(
            texttemplate="%{text:.1f}%",
            textposition="inside",
        )

        fig_det.update_layout(
            yaxis_title="Detection %",
            xaxis_title="Assay",
        )

        fig_det.update_yaxes(range=[0, 110])

        st.plotly_chart(fig_det, use_container_width=True)

        output = BytesIO()

        assay_layout_df = pd.DataFrame(st.session_state.qc_assay_layout)

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            metadata_df.to_excel(writer, sheet_name="LOT_information", index=False)
            assay_layout_df.to_excel(writer, sheet_name="Assay_layout", index=False, header=False)
            qc_assessment.to_excel(writer, sheet_name="QC_assessment", index=False)
            flat_ch2.to_excel(writer, sheet_name="CH2_summary", index=False)
            flat_ch3.to_excel(writer, sheet_name="CH3_summary", index=False)
            full_df.to_excel(writer, sheet_name="Full_Data_Processed", index=False)

        st.download_button(
            "Download QC Excel report",
            data=output.getvalue(),
            file_name="qc_pod_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


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

        for raw_df in [ch2, ch3]:
            raw_df["Loaded_num"] = (
                raw_df["Loaded"]
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

        box_plot_df = ch2 if channel == "CH2" else ch3
        detection_plot_df = flat_ch2 if channel == "CH2" else flat_ch3

        metric_options = {
            "Cq": "Cq",
            "Amplitude": "Ampl",
            "Slope": "Slope",
            "Background": "Background",
        }

        metric_label = st.selectbox(
            "Select metric",
            list(metric_options.keys()),
            key="within_metric"
        )

        metric_column = metric_options[metric_label]

        if metric_column not in box_plot_df.columns:
            st.warning(f"No column found for {metric_label}. Available columns: {', '.join(box_plot_df.columns)}")
        else:
            box_plot_df[metric_column] = pd.to_numeric(
                box_plot_df[metric_column],
                errors="coerce"
            )

            loaded_order = (
                box_plot_df
                .sort_values(["Loaded_num", "Loaded"], ascending=[False, True])
                ["Loaded"]
                .dropna()
                .unique()
            )

            fig = px.box(
                box_plot_df,
                x="Loaded",
                y=metric_column,
                points="all",
                title=f"{metric_label} by Loaded ({channel})",
                category_orders={"Loaded": loaded_order},
            )

            fig.update_layout(
                xaxis_title="Loaded",
                yaxis_title=metric_label,
                showlegend=False,
            )

            st.plotly_chart(fig, use_container_width=True)

        st.header("Detection rate")

        fig_det = px.bar(
            detection_plot_df,
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