import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import os

# -----------------------------
# Load your dataframe
# -----------------------------
# Replace with your file path
file_path = "/Users/lenahirzel/Desktop/Diaxxo_temp_save/Plotting_analysis/Pod-to-pod.csv"

df = pd.read_csv(file_path, sep=";")

# Optional: check columns
print(df.columns)
print(df.head())
print(df.head(10))

base_dir = os.path.dirname(os.path.abspath(file_path))
plot_dir = os.path.join(base_dir, "plots")

os.makedirs(plot_dir, exist_ok=True)

# Make sure numeric columns are numeric
for col in ["Well ID", "Cq", "Ampl."]:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# Optional: replace invalid Cq values (-1) with NaN for plotting
df["Cq_plot"] = df["Cq"].replace(-1, pd.NA)

# -----------------------------
# Plot settings
# -----------------------------
sns.set(style="whitegrid")

metrics = {
    "Cq_plot": "Cq",
    "Ampl.": "Ampl."
}

channels = ["CH2", "CH3"]

# -----------------------------
# Create plots
# -----------------------------
for channel in channels:
    for metric in metrics:

        df_ch = df[df["Channel"] == channel].copy()

        # ax = sns.boxplot(
        #     data=df_ch,
        #     x="Well ID",
        #     y=metric,
        #     hue="Condition"
        # )
        g = sns.catplot(
            data=df_ch,
            x="Well ID",
            y=metric,
            hue="Condition",
            col="Loaded",
            kind="strip",
            jitter=True,
            dodge=True
        )

        g.fig.suptitle(f"{metric} by Well Position ({channel})", y=1.05)
        #ax.set_title(f"{metric} by Well Position ({channel})")
        # save
        out_png = f"{plot_dir}/well_dis_{channel}_{metric}.png"
        out_pdf = f"{plot_dir}/well_dis_{channel}_{metric}.pdf"

        plt.savefig(out_png, dpi=300, bbox_inches="tight")
        plt.savefig(out_pdf, bbox_inches="tight")

        plt.show()

    # fig, axes = plt.subplots(
    #     nrows=2,
    #     ncols=1,
    #     figsize=(14, 10),
    #     sharex=True
    # )
    #
    # for ax, (metric_col, metric_name) in zip(axes, metrics.items()):
    #
    #     sns.scatterplot(
    #         data=df_ch,
    #         x="Well ID",
    #         y=metric_col,
    #         hue="Condition",
    #         style="Loaded",
    #         s=120,
    #         ax=ax
    #     )
    #
    #     # connect V2/V3 mean values by well
    #     mean_df = (
    #         df_ch.groupby(["Well ID", "Condition"])[metric_col]
    #         .mean()
    #         .reset_index()
    #     )
    #
    #     sns.lineplot(
    #         data=mean_df,
    #         x="Well ID",
    #         y=metric_col,
    #         hue="Condition",
    #         marker="o",
    #         legend=False,
    #         ax=ax
    #     )
    #
    #     ax.set_title(f"{metric_name} by Well Position ({channel})")
    #     ax.set_xlabel("Well ID")
    #     ax.set_ylabel(metric_name)
    #
    # plt.tight_layout()
    # plt.show()