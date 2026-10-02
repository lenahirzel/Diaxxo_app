from io import BytesIO, StringIO
import csv
from pathlib import Path
from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import streamlit as st
import pandas as pd
from analysis_v7 import run_analysis
from qc_analysis import (
    add_qc_sample_and_assay_columns,
    assay_layout_to_text,
    assess_qc_results,
    build_combined_layout_lines,
    expectations_to_dataframe,
    get_product_config,
    get_qc_config_path,
    load_qc_config,
)
from pod_to_pod_comparison_v2 import (
    run_pod_to_pod_comparison,
    figure_to_png_bytes,
    figure_to_pdf_bytes,
)
import plotly.express as px

QC_WORD_TEMPLATE_PATH = (
    Path(__file__).resolve().parent
    / "templates"
    / "SN8491_CoA-diaxxoPod-200_v1.docx"
)

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

def format_report_value(value):
    """Format values safely for the Word report."""
    if pd.isna(value):
        return ""

    return str(value)


def set_cell_background(cell, hex_color):
    """Set Word table cell background color."""
    cell_properties = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), hex_color)
    cell_properties.append(shading)


def replace_text_in_paragraph(paragraph, replacements):
    """Replace placeholders inside a paragraph while preserving basic Word styling."""
    full_text = "".join(run.text for run in paragraph.runs)

    if not any(placeholder in full_text for placeholder in replacements):
        return

    for placeholder, replacement in replacements.items():
        full_text = full_text.replace(placeholder, str(replacement))

    for run in paragraph.runs:
        run.text = ""

    if paragraph.runs:
        paragraph.runs[0].text = full_text
    else:
        paragraph.add_run(full_text)


def replace_text_everywhere(document, replacements):
    """Replace placeholders in normal text, tables, headers, and footers."""
    for paragraph in document.paragraphs:
        replace_text_in_paragraph(paragraph, replacements)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text_in_paragraph(paragraph, replacements)

    for section in document.sections:
        for header_footer in [
            section.header,
            section.first_page_header,
            section.even_page_header,
            section.footer,
            section.first_page_footer,
            section.even_page_footer,
        ]:
            for paragraph in header_footer.paragraphs:
                replace_text_in_paragraph(paragraph, replacements)

            for table in header_footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            replace_text_in_paragraph(paragraph, replacements)


def insert_dataframe_after_paragraph(paragraph, dataframe):
    """Insert QC dataframe as a Word table directly after the given paragraph."""
    parent = paragraph._parent
    column_count = len(dataframe.columns)

    try:
        table = parent.add_table(rows=1, cols=column_count)
    except TypeError:
        table = parent.add_table(rows=1, cols=column_count, width=Inches(6.5))

    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    paragraph._p.addnext(table._tbl)

    header_cells = table.rows[0].cells

    for column_index, column_name in enumerate(dataframe.columns):
        header_cells[column_index].text = str(column_name)
        set_cell_background(header_cells[column_index], "D9EAF7")

        for header_paragraph in header_cells[column_index].paragraphs:
            for run in header_paragraph.runs:
                run.bold = True

    for _, dataframe_row in dataframe.iterrows():
        row_cells = table.add_row().cells

        qc_assessment = str(dataframe_row.get("QC assessment", "")).lower()

        if qc_assessment == "passed":
            row_color = "D4EDDA"
        else:
            row_color = "F8D7DA"

        for column_index, column_name in enumerate(dataframe.columns):
            row_cells[column_index].text = format_report_value(dataframe_row[column_name])
            row_cells[column_index].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_background(row_cells[column_index], row_color)

    return table
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    paragraph._p.addnext(table._tbl)

    header_cells = table.rows[0].cells

    for column_index, column_name in enumerate(dataframe.columns):
        header_cells[column_index].text = str(column_name)
        set_cell_background(header_cells[column_index], "D9EAF7")

        for header_paragraph in header_cells[column_index].paragraphs:
            for run in header_paragraph.runs:
                run.bold = True

    for _, dataframe_row in dataframe.iterrows():
        row_cells = table.add_row().cells

        qc_assessment = str(dataframe_row.get("QC assessment", "")).lower()

        if qc_assessment == "passed":
            row_color = "D4EDDA"
        else:
            row_color = "F8D7DA"

        for column_index, column_name in enumerate(dataframe.columns):
            row_cells[column_index].text = format_report_value(dataframe_row[column_name])
            row_cells[column_index].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_background(row_cells[column_index], row_color)

    return table


def replace_qc_table_placeholder(document, qc_assessment):
    """Replace {{QC_TABLE}} placeholder with the QC assessment table."""
    for paragraph in document.paragraphs:
        if "{{QC_TABLE}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{QC_TABLE}}", "")
            insert_dataframe_after_paragraph(paragraph, qc_assessment)
            return

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if "{{QC_TABLE}}" in paragraph.text:
                        paragraph.text = paragraph.text.replace("{{QC_TABLE}}", "")
                        insert_dataframe_after_paragraph(paragraph, qc_assessment)
                        return


def create_qc_word_report_from_template(qc_metadata, qc_assessment, overall_qc_passed):
    """Create a QC Word report by filling an existing .docx template."""
    if not QC_WORD_TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Word template was not found: {QC_WORD_TEMPLATE_PATH}"
        )

    document = Document(QC_WORD_TEMPLATE_PATH)

    replacements = {
        "{{PRODUCT_NUMBER}}": format_report_value(qc_metadata.get("Product number")),
        "{{LOT_SN}}": format_report_value(qc_metadata.get("LOT serial number")),
        "{{MANUFACTURING_DATE}}": format_report_value(qc_metadata.get("Manufacturing date")),
        "{{EXPIRATION_DATE}}": format_report_value(qc_metadata.get("Expiration date")),
        "{{QC_RESULT}}": "PASSED" if overall_qc_passed else "NOT PASSED",
    }

    replace_text_everywhere(document, replacements)
    replace_qc_table_placeholder(document, qc_assessment)

    output = BytesIO()
    document.save(output)
    output.seek(0)

    return output.getvalue()

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

    st.info(
        "QC configuration is loaded from qc_product_config.json in the repository. "
        "To make permanent changes, edit that JSON file and commit it to the repo."
    )

    st.markdown(
        "This assay layout is linked to the product number. "
        "It defines which assay is present in each pod position."
    )

    if assay_layout:
        st.text_area(
            "Assay loading scheme",
            value=assay_layout_to_text(assay_layout),
            height=160,
            disabled=True,
            key=f"assay_layout_{product_number}",
        )
    else:
        st.warning(
            f"No assay loading scheme found for product `{product_number}` "
            "in qc_product_config.json."
        )

    st.divider()

    st.subheader(f"Product {product_number} QC expectations")

    st.markdown(
        "These expectations are specific to the combination of "
        "**product number + QC sample + assay + channel**."
    )

    expectations_df = expectations_to_dataframe(qc_expectations)

    if expectations_df.empty:
        st.warning(
            f"No QC expectations found for product `{product_number}` "
            "in qc_product_config.json."
        )
    else:
        st.dataframe(expectations_df, use_container_width=True)

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

        required_box_plot_columns = ["Assay", "QC_sample", metric_column]
        missing_box_plot_columns = [
            column
            for column in required_box_plot_columns
            if column not in box_plot_df.columns
        ]

        if box_plot_df.empty:
            st.warning(f"No {channel} data available for the box plot.")

        elif missing_box_plot_columns:
            st.warning(
                "Cannot create the QC box plot because the following column(s) are missing: "
                + ", ".join(missing_box_plot_columns)
                + f". Available columns: {', '.join(box_plot_df.columns)}"
            )

        else:
            box_plot_df = box_plot_df.copy()
            box_plot_df[metric_column] = pd.to_numeric(
                box_plot_df[metric_column],
                errors="coerce",
            )

            box_plot_df = box_plot_df.dropna(
                subset=["Assay", "QC_sample", metric_column]
            )

            if box_plot_df.empty:
                st.warning(
                    f"No plottable {metric_label} values available for {channel}."
                )
            else:
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

        required_detection_columns = ["Assay", "QC_sample", "QC_Detection_%"]
        missing_detection_columns = [
            column
            for column in required_detection_columns
            if column not in detection_plot_df.columns
        ]

        if detection_plot_df.empty:
            st.warning(f"No {channel} data available for the detection plot.")

        elif missing_detection_columns:
            st.warning(
                "Cannot create the detection plot because the following column(s) are missing: "
                + ", ".join(missing_detection_columns)
                + f". Available columns: {', '.join(detection_plot_df.columns)}"
            )

        else:
            detection_plot_df = detection_plot_df.copy()
            detection_plot_df["QC_Detection_%"] = pd.to_numeric(
                detection_plot_df["QC_Detection_%"],
                errors="coerce",
            )

            detection_plot_df = detection_plot_df.dropna(
                subset=["Assay", "QC_sample", "QC_Detection_%"]
            )

            if detection_plot_df.empty:
                st.warning(f"No plottable detection data available for {channel}.")

            else:
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

        try:
            word_report = create_qc_word_report_from_template(
                qc_metadata=qc_metadata,
                qc_assessment=qc_assessment,
                overall_qc_passed=overall_qc_passed,
            )

            st.download_button(
                "Download QC Word report",
                data=word_report,
                file_name=f"qc_pod_report_product_{product_number}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        except FileNotFoundError as error:
            st.warning(str(error))


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

        required_box_plot_columns = ["Loaded", metric_column]
        missing_box_plot_columns = [
            column
            for column in required_box_plot_columns
            if column not in box_plot_df.columns
        ]

        if box_plot_df.empty:
            st.warning(f"No {channel} data available for the box plot.")

        elif missing_box_plot_columns:
            st.warning(
                "Cannot create the QC box plot because the following column(s) are missing: "
                + ", ".join(missing_box_plot_columns)
                + f". Available columns: {', '.join(box_plot_df.columns)}"
            )

        else:
            box_plot_df = box_plot_df.copy()
            box_plot_df[metric_column] = pd.to_numeric(
                box_plot_df[metric_column],
                errors="coerce",
            )

            box_plot_df = box_plot_df.dropna(
                subset=["Loaded", metric_column]
            )

            if box_plot_df.empty:
                st.warning(
                    f"No plottable {metric_label} values available for {channel}."
                )
            else:
                fig = px.box(
                    box_plot_df,
                    x="Loaded",
                    y=metric_column,
                    points="all",
                    title=f"{metric_label} by Loaded ({channel})",
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