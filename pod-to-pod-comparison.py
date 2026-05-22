import os
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# ---------------------------
# SETTINGS
# ---------------------------

sns.set_style("whitegrid")
sns.set_context("talk")  # publication-like scaling

# ---------------------------
# 1. LOAD CSV
# ---------------------------

# Replace with your file path
file_path = "/Users/lenahirzel/Desktop/Diaxxo_temp_save/Plotting_analysis/Pod-to-pod.csv"

df = pd.read_csv(file_path, sep=";")

# Optional: check columns
print(df.info)
print("\n--- DATAFRAME SHAPE ---")
print(f"Rows: {df.shape[0]}")
print(f"Columns: {df.shape[1]}")
print(df.columns)
print(df.head())
print(df.head(10))

base_dir = os.path.dirname(os.path.abspath(file_path))
plot_dir = os.path.join(base_dir, "plots")

os.makedirs(plot_dir, exist_ok=True)


# ---------------------------
# CLEANING
# ---------------------------
def prepare_data(df):
    df = df.copy()

    for col in ["Cq", "Ampl.", "Slope"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["Detected"] = df["Cq"] > 0

    df_valid = df[df["Cq"] > 0].copy()

    return df, df_valid



# ---------------------------
# MAIN FIGURE (BY CHANNEL)
# ---------------------------
def make_publication_figure(df_valid):

    metrics = ["Cq", "Ampl.", "Slope"]
    channels = df_valid["Channel"].unique()

    hue_order = sorted(df_valid["Condition"].dropna().unique())
    palette = dict(zip(hue_order, sns.color_palette("tab10", n_colors=len(hue_order))))

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

            # remove subplot legends
            if ax.get_legend() is not None:
                ax.legend_.remove()

        # shared legend
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

        plt.tight_layout(rect=[0, 0, 0.88, 1])

        # save
        out_png = f"{plot_dir}/qpcr_{ch}.png"
        out_pdf = f"{plot_dir}/qpcr_{ch}.pdf"

        plt.savefig(out_png, dpi=300, bbox_inches="tight")
        plt.savefig(out_pdf, bbox_inches="tight")

        plt.show()


# ---------------------------
# DETECTION PLOT (SEPARATE FIGURE)
# ---------------------------
def plot_detection(df_all):
    df_ch3 = df_all[df_all["Channel"] == "CH3"].copy()

    fig, ax = plt.subplots(figsize=(15, 8))

    sns.barplot(
        data=df_ch3,
        x="Loaded",
        y="Detected",
        hue="Condition",
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

    # save
    out_png = f"{plot_dir}/detection_rate.png"
    out_pdf = f"{plot_dir}/detection_rate.pdf"

    plt.savefig(out_png, dpi=300, bbox_inches="tight")
    plt.savefig(out_pdf, bbox_inches="tight")

    plt.show()


# ---------------------------
# RUN
# ---------------------------
df_all, df_valid = prepare_data(df)

print(df_all.groupby(["Condition", "Channel", "Loaded"])["Detected"].mean())

make_publication_figure(df_valid)
#make_publication_figure(df_all)
plot_detection(df_all)