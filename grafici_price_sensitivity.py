from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter

# --- PATHS ---
excel_path = Path(
    r"C:\CityEnergyAnalyst\Paper_prova\Elaboration_REC\simulations 0.2 PV\comparison.xlsx"
)
output_dir = Path(
    r"C:\CityEnergyAnalyst\Paper_prova\Elaboration_REC\simulations 0.2 PV"
)

# --- STYLE ---
TITLE_FS = 10
LABEL_FS = 9
TICK_FS = 8
LEGEND_FS = 8
GRID_COLOR = "#d9d9d9"

MARKER = "o"
MARKER_SIZE = 4

def load_three_scenarios(sheet_name: str):
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    idx_sc1 = df[df.iloc[:, 0] == "Scenario I"].index[0]
    idx_sc2 = df[df.iloc[:, 0] == "Scenario II"].index[0]

    df_b = df.iloc[0:idx_sc1].copy().dropna(subset=["Baseline"])
    df_s1 = df.iloc[idx_sc1 + 1:idx_sc2].copy().dropna(subset=["Baseline"])
    df_s2 = df.iloc[idx_sc2 + 1:].copy().dropna(subset=["Baseline"])

    return df_b, df_s1, df_s2


def plot_one_axis(ax, sheet_name: str, title: str, colors):
    df_b, df_s1, df_s2 = load_three_scenarios(sheet_name)

    x = df_b["Baseline"]

    # € -> k€
    y_b = df_b["Total net costs"] / 1000.0
    y_s1 = df_s1["Total net costs"] / 1000.0
    y_s2 = df_s2["Total net costs"] / 1000.0

    ax.plot(
        x, y_b,
        label="Baseline",
        color=colors[0],
        linewidth=2,
        marker=MARKER,
        markersize=MARKER_SIZE
    )
    ax.plot(
        x, y_s1,
        label="Scenario I",
        color=colors[1],
        linewidth=2,
        marker=MARKER,
        markersize=MARKER_SIZE
    )
    ax.plot(
        x, y_s2,
        label="Scenario II",
        color=colors[2],
        linewidth=2,
        marker=MARKER,
        markersize=MARKER_SIZE
    )

    ax.set_title(title, fontsize=TITLE_FS)
    ax.set_xlabel("Tariff variation [%]", fontsize=LABEL_FS)
    ax.set_ylabel("Total net costs [k€]", fontsize=LABEL_FS)

    ax.grid(True, color=GRID_COLOR, linewidth=0.6)
    ax.set_ylim(bottom=0)

    ax.tick_params(axis="both", labelsize=TICK_FS)
    ax.xaxis.set_major_formatter(PercentFormatter(xmax=1))

    # legenda locale, in alto fuori dal grafico
    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, 1.15),
        ncol=3,
        fontsize=LEGEND_FS,
        frameon=False
    )


# ------------------------------------------------------------
# COLORS
# ------------------------------------------------------------
green_colors = ["#0b3d0b", "#2e8b57", "#9ad29a"]
purple_colors = ["#3b0054", "#7a1fa2", "#c39bd3"]

# ------------------------------------------------------------
# FIGURE
# ------------------------------------------------------------
fig, axes = plt.subplots(1, 2, figsize=(12, 6))

plot_one_axis(
    axes[0],
    sheet_name="variation electricity",
    title="Electricity price variation",
    colors=green_colors
)

plot_one_axis(
    axes[1],
    sheet_name="variation NG",
    title="Natural Gas price variation",
    colors=purple_colors
)

fig.tight_layout(rect=[0, 0, 1, 0.90])

# --- SAVE ---
out_file = output_dir / "sensitivity_electricity_vs_NG_kEUR.jpg"
fig.savefig(out_file, dpi=400, bbox_inches="tight")

plt.show()

