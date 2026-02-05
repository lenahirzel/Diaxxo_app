import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
from scipy.stats import ttest_ind
import zipfile
import os

# --- Helpers ---

def add_replicate_count(summary_df, df):
    n_loaded = df.groupby("Loaded")["Cq"].count()
    summary_df["QC_N_loaded"] = n_loaded
    return summary_df

def flatten_summary(summary_df, channel):
    flat = summary_df.copy()
    flat.columns = ["_".join(col).strip() if isinstance(col, tuple) else col for col in flat.columns.values]
    flat["Channel"] = channel
    return flat.reset_index()

def run_analysis_single(df, layout_lines):
    # Clean columns
    df.columns = df.columns.str.strip().str.replace(" ", "_")
    df["Sample_Name"] = df["Sample_Name"].ffill()
    df["Well_ID"] = df["Well_ID"].ffill()

    # Parse layout
    layout = [row.split("\t") for row in layout_lines]
    mapping = {}
    well_number = 1
    for r in range(len(layout)):
        for c in range(len(layout[0])):
            mapping[well_number] = layout[r][c]
            well_number += 1
    df["Loaded"] = df["Well_ID"].map(mapping)

    # Split channels
    ch2 = df[df["Channel"]=="CH2"].copy()
    ch3 = df[df["Channel"]=="CH3"].copy()

    ch2.loc[ch2["Cq"]==-1, "Cq"] = np.nan
    ch3.loc[ch3["Cq"]==-1, "Cq"] = np.nan

    # Summaries
    summary_ch2 = ch2.groupby("Loaded")[["Cq","Ampl.","Slope"]].agg(["mean","std"])
    summary_ch3 = ch3.groupby("Loaded")[["Cq","Ampl.","Slope"]].agg(["mean","std"])

    summary_ch2 = add_replicate_count(summary_ch2, ch2)
    summary_ch3 = add_replicate_count(summary_ch3, ch3)

    # Detection %
    summary_ch2["Detection_%_"] = ch2.groupby("Loaded")["Classification"].apply(lambda x: (x=="POSITIVE").mean()*100).values
    summary_ch3["Detection_%_"] = ch3.groupby("Loaded")["Classification"].apply(lambda x: (x=="POSITIVE").mean()*100).values

    flat_ch2 = flatten_summary(summary_ch2, "CH2")
    flat_ch3 = flatten_summary(summary_ch3, "CH3")

    return df, summary_ch2, summary_ch3, flat_ch2, flat_ch3, ch2, ch3

def parse_multi_experiment_excel(uploaded_file):
    raw = pd.read_excel(uploaded_file, header=None)

    # find experiment starts
    id_rows = raw.index[
        raw[0].astype(str).str.startswith("ID:")
    ].tolist()
    id_rows.append(len(raw))

    experiments = []

    for i in range(len(id_rows) - 1):
        block = raw.iloc[id_rows[i]:id_rows[i+1]].reset_index(drop=True)

        exp_id = block.iloc[0, 0].replace("ID:", "").strip()
        exp_name = block.loc[block[0].str.startswith("Name:"), 0].iloc[0].replace("Name:", "").strip()
        device = block.loc[block[0].str.startswith("Device:"), 0].iloc[0].replace("Device:", "").strip()

        header_row = block.index[block[0] == "Sample Name"][0]

        df = block.iloc[header_row + 1:].copy()
        df.columns = block.iloc[header_row]
        df = df.dropna(how="all")

        df["Sample Name"] = df["Sample Name"].ffill()
        df["Well ID"] = df["Well ID"].ffill()

        df["Experiment_ID"] = exp_id
        df["Experiment_Name"] = exp_name
        df["Device"] = device

        experiments.append(df)

    df_all = pd.concat(experiments, ignore_index=True)

    # clean numerics
    for col in ["Cq", "Ampl.", "Slope"]:
        df_all[col] = pd.to_numeric(df_all[col], errors="coerce")

    df_all["Detected"] = df_all["Classification"] == "POSITIVE"

    return df_all


def summarize_multi_experiment(df):
    numeric_cols = ["Cq","Ampl.","Slope"]
    summaries = []
    for exp_id, exp_df in df.groupby("Experiment_ID"):
        for channel, ch_df in exp_df.groupby("Channel"):
            summary = ch_df.groupby("Loaded")[numeric_cols].agg(["mean","std"]).reset_index()
            summary["Experiment_ID"] = exp_id
            summary["Experiment_Name"] = ch_df["Experiment_Name"].iloc[0]
            summary["Channel"] = channel
            summary["Detection_%_"] = ch_df.groupby("Loaded")["Classification"].apply(lambda x: (x=="POSITIVE").mean()*100).values
            summary["N_replicates"] = ch_df.groupby("Loaded")["Cq"].count().values
            summaries.append(summary)
    combined_summary = pd.concat(summaries, ignore_index=True)
    return combined_summary

# --- Streamlit app ---

st.title("Experiment analysis App")

mode = st.sidebar.radio("Select mode:", ["Single pod", "Multi-experiment"])

# --- SINGLE POD ---
if mode=="Single pod":
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
    layout_text = st.text_area(
        "Paste pod loading scheme (tab-separated)",
        height=200,
        placeholder="T1\tT1\tT2\tT2\nT3\tT3\tT4\tT4"
    )

    if uploaded_file and layout_text:
        if st.button("Run pod analysis"):
            df = pd.read_excel(uploaded_file)
            layout_lines = layout_text.strip().split("\n")
            results = run_analysis_single(df, layout_lines)
            (
                st.session_state.full_df,
                st.session_state.ch2,
                st.session_state.ch3,
                st.session_state.flat_ch2,
                st.session_state.flat_ch3,
                st.session_state.raw_ch2,
                st.session_state.raw_ch3,
            ) = results

            st.success("Analysis completed!")

            st.subheader("CH2 Summary")
            st.dataframe(st.session_state.flat_ch2)

            st.subheader("CH3 Summary")
            st.dataframe(st.session_state.flat_ch3)

            # --- Plots ---
            figures = {}

            # CH2 boxplots
            for metric in ["Cq","Ampl.","Slope"]:
                fig = px.box(
                    st.session_state.raw_ch2,
                    x="Loaded",
                    y=metric,
                    points="all",
                    title=f"{metric} by Loaded (CH2)"
                )
                figures[f"CH2_{metric}_box"] = fig

            # CH3 boxplots
            for metric in ["Cq","Ampl.","Slope"]:
                fig = px.box(
                    st.session_state.raw_ch3,
                    x="Loaded",
                    y=metric,
                    points="all",
                    title=f"{metric} by Loaded (CH3)"
                )
                figures[f"CH3_{metric}_box"] = fig

            # Detection rate
            for ch, flat in zip(["CH2","CH3"], [st.session_state.flat_ch2, st.session_state.flat_ch3]):
                fig_det = px.bar(
                    flat,
                    x="Loaded",
                    y="Detection_%__",
                    text="Detection_%__",
                    title=f"Detection rate ({ch})",
                    range_y=[0,100]
                )
                fig_det.update_traces(
                    text=flat["Detection_%__"].round(1).astype(str)+"%",
                    textposition="inside"
                )
                for i,row in flat.iterrows():
                    fig_det.add_annotation(
                        x=row["Loaded"],
                        y=row["Detection_%__"],
                        text=f"n={int(row['QC_N_loaded_'])}",
                        showarrow=False,
                        yshift=12
                    )
                figures[f"{ch}_detection"] = fig_det

            # --- Show plots ---
            st.sidebar.header("Plots")
            show_ch2 = st.sidebar.checkbox("Show CH2 plots", True)
            show_ch3 = st.sidebar.checkbox("Show CH3 plots", True)
            show_detection = st.sidebar.checkbox("Show detection rate", True)

            if show_ch2:
                st.subheader("CH2 plots")
                for key, fig in figures.items():
                    if key.startswith("CH2_") and "detection" not in key:
                        st.plotly_chart(fig, use_container_width=True)
            if show_ch3:
                st.subheader("CH3 plots")
                for key, fig in figures.items():
                    if key.startswith("CH3_") and "detection" not in key:
                        st.plotly_chart(fig, use_container_width=True)
            if show_detection:
                st.subheader("Detection rate")
                st.plotly_chart(figures["CH2_detection"], use_container_width=True)
                st.plotly_chart(figures["CH3_detection"], use_container_width=True)

# --- MULTI-EXPERIMENT ---
elif mode=="Multi-experiment":
    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx"]
    )
    if uploaded_file:
        if st.button("Run multi-experiment analysis"):
            full_df = parse_multi_experiment_excel(uploaded_file)
            combined_summary = summarize_multi_experiment(df_multi)

            st.success("Multi-experiment summary completed!")
            st.dataframe(combined_summary)

            # --- Plots ---
            figures = {}
            for metric in ["Cq","Ampl.","Slope"]:
                fig = px.box(
                    df_multi,
                    x="Loaded",
                    y=metric,
                    color="Experiment_ID",
                    facet_col="Channel",
                    points="all",
                    title=f"{metric} by Loaded across experiments"
                )
                figures[f"{metric}_box"] = fig
            # Detection
            fig_det = px.bar(
                combined_summary,
                x="Loaded",
                y="Detection_%_",
                color="Experiment_ID",
                facet_col="Channel",
                text="Detection_%_",
                title="Detection rate across experiments",
                range_y=[0,100]
            )
            for i,row in combined_summary.iterrows():
                fig_det.add_annotation(
                    x=row["Loaded"],
                    y=row["Detection_%_"],
                    text=f"n={int(row['N_replicates'])}",
                    showarrow=False,
                    yshift=10
                )
            figures["detection"] = fig_det

            # --- Show plots ---
            st.subheader("Plots")
            for key, fig in figures.items():
                st.plotly_chart(fig, use_container_width=True)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        full_df.to_excel(writer, sheet_name="Full_Data", index=False)