from io import BytesIO

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt


# ---------------------------
# SETTINGS
# ---------------------------

sns.set_style("whitegrid")
sns.set_context("talk")


# ---------------------------
# LOAD CSV
# ---------------------------

def load_pod_to_pod_csv(uploaded_file, sep=";"):
    df = pd.read_csv(uploaded_file, sep=sep)
    df.columns = df.columns.str.strip()
    return df


# ---------------------------
# CLEANING
# ---------------------------

def prepare_data(df):
    df = df.copy()
    df.columns = df.columns.str.strip()

    required_columns = ["Cq", "Ampl", "Slope", "Channel", "Condition", "Loaded"]
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        raise ValueError(
            "Missing required columns: "
            + ", ".join(missing_columns)
        )

    for col in ["Cq", "Ampl", "Slope"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["Detected"] = df["Cq"] > 0

    df_valid = df[df["Cq"] > 0].copy()

    return df, df_valid


# ---------------------------
# MAIN FIGURE BY CHANNEL
# ---------------------------

def make_publication_figures(df_valid):
    figures = {}

    metrics = ["Cq", "Ampl", "Slope"]
    channels = df_valid["Channel"].dropna().unique()

    hue_order = sorted(df_valid["Condition"].dropna().unique())
    palette = dict(
        zip(
            hue_order,
            sns.color_palette("tab10", n_colors=len(hue_order))
        )
    )

    for ch in channels:
        df_ch = df_valid[df_valid["Channel"] == ch]

        fig, axes = plt.subplots(
            nrows=1,
            ncols=3,
            figsize=(20, 8),
            sharex=True
        )

        for ax, metric in zip(axes, metrics):
            sns.boxplot(
                data=df_ch,
                x="Loaded",
                y=metric,
                hue="Condition",
                hue_order=hue_order,
                palette=palette,
                ax=ax,
                showfliers=False
            )

            sns.stripplot(
                data=df_ch,
                x="Loaded",
                y=metric,
                hue="Condition",
                hue_order=hue_order,
                palette=palette,
                dodge=True,
                alpha=0.5,
                linewidth=0,
                ax=ax
            )

            ax.set_title(f"{metric} ({ch})")
            ax.set_xlabel("Loaded")

            if ax.get_legend() is not None:
                ax.legend_.remove()

        handles, labels = axes[0].get_legend_handles_labels()
        n = df_ch["Condition"].nunique()

        fig.legend(
            handles[:n],
            labels[:n],
            title="Condition",
            loc="center left",
            bbox_to_anchor=(0.85, 0.5),
            ncol=1,
            fontsize=14,
            title_fontsize=14
        )

        fig.tight_layout(rect=[0, 0, 0.88, 1])

        figures[ch] = fig

    return figures


# ---------------------------
# DETECTION PLOT
# ---------------------------

def make_detection_figure(df_all):
    df_ch3 = df_all[df_all["Channel"] == "CH3"].copy()

    if df_ch3.empty:
        return None

    hue_order = sorted(df_all["Condition"].dropna().unique())
    palette = dict(
        zip(
            hue_order,
            sns.color_palette("tab10", n_colors=len(hue_order))
        )
    )

    fig, ax = plt.subplots(figsize=(15, 8))

    sns.barplot(
        data=df_ch3,
        x="Loaded",
        y="Detected",
        hue="Condition",
        hue_order=hue_order,
        palette=palette,
        ax=ax
    )

    ax.set_ylabel("Detection rate")
    ax.set_ylim(0, 1)
    ax.set_title("Detection rate by condition (CH3 only)")

    ax.legend(
        title="Condition",
        loc="center left",
        bbox_to_anchor=(0.9, 0.5),
        fontsize=14,
        title_fontsize=14,
        frameon=False
    )

    fig.tight_layout()

    return fig


# ---------------------------
# EXPORT HELPERS
# ---------------------------

def figure_to_png_bytes(fig):
    output = BytesIO()
    fig.savefig(output, format="png", dpi=300, bbox_inches="tight")
    output.seek(0)
    return output.getvalue()


def figure_to_pdf_bytes(fig):
    output = BytesIO()
    fig.savefig(output, format="pdf", bbox_inches="tight")
    output.seek(0)
    return output.getvalue()


# ---------------------------
# RUN FROM STREAMLIT APP
# ---------------------------

def run_pod_to_pod_comparison(data):
    if isinstance(data, pd.DataFrame):
        df = data.copy()
    else:
        df = load_pod_to_pod_csv(data)

    df_all, df_valid = prepare_data(df)

    summary = (
        df_all
        .groupby(["Condition", "Channel", "Loaded"], dropna=False)["Detected"]
        .mean()
        .reset_index()
        .rename(columns={"Detected": "Detection_rate"})
    )

    publication_figures = make_publication_figures(df_valid)
    detection_figure = make_detection_figure(df_all)

    return {
        "df_all": df_all,
        "df_valid": df_valid,
        "summary": summary,
        "publication_figures": publication_figures,
        "detection_figure": detection_figure,
    }