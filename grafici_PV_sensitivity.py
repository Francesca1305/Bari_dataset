from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# --- INPUT ---
excel_path = Path(r"C:\CityEnergyAnalyst\Paper_prova\Elaboration_REC\sensitivity PV\comparison.xlsx")
sheets = ["Baseline", "Scenario I", "Scenario II"]

# Mapping efficienze PV
eff_map = {1: 18, 5: 20, 6: 22, 7: 24, 8: 26}

# --- LOAD & PREP ---
dfs = {}
for sh in sheets:
    df = pd.read_excel(excel_path, sheet_name=sh)

    df["PV_num"] = df["Scenario"].astype(str).str.extract(r"PV(\d+)").astype(int)
    df = df.sort_values("PV_num").reset_index(drop=True)
    df["Eff_pct"] = df["PV_num"].map(eff_map).astype(float)

    # Conversioni unità
    df["total_SC_MWh"] = df["total_SC_sum"] / 1000.0
    df["CSC_MWh"] = df["CSC_sum"] / 1000.0
    df["Net_Electricity_Costs_kEUR"] = (
        df["Energy costs REC_sum"] - df["Energy revenues total_sum"]
    ) / 1000.0

    dfs[sh] = df

# Asse x: efficienze
x = dfs["Baseline"]["Eff_pct"].values
x_labels = [f"{int(v)}%" for v in x]

# Colori per metrica (BAU → Scenario I → Scenario II)
colors_by_plot = {
    "total_SC_MWh": {
        "Baseline": "darkred",
        "Scenario I": "red",
        "Scenario II": "darkorange"
    },
    "CSC_MWh": {
        "Baseline": "saddlebrown",
        "Scenario I": "peru",
        "Scenario II": "gold"
    },
    "Net_Electricity_Costs_kEUR": {
        "Baseline": "navy",
        "Scenario I": "royalblue",
        "Scenario II": "lightblue"
    },
}

plots = [
    ("total_SC_MWh", "Physical self-consumption", "MWh/y"),
    ("CSC_MWh", "Collective self-consumption (CSC)", "MWh/y"),
    ("Net_Electricity_Costs_kEUR", "REC Net Electricity Costs (costs − revenues)", "k€/y"),
]

# --- PLOT ---
fig, axes = plt.subplots(1, 3, figsize=(12, 4))
axes = np.array(axes).flatten()

for ax, (col, title, unit) in zip(axes, plots):

    for sh in sheets:
        df = dfs[sh]
        ax.plot(
            df["Eff_pct"].values,
            df[col].values,
            marker="o",
            linewidth=1.5,
            color=colors_by_plot[col][sh],
            label=sh
        )

    ax.set_title(title, fontsize=10)
    ax.set_xlabel("PV module efficiency", fontsize=10)
    ax.set_ylabel(unit, fontsize=8)
    ax.tick_params(axis="both", labelsize=8)

    ax.grid(True, which="major", linestyle="-", linewidth=0.6, alpha=0.6)
    ax.set_xticks(x)
    ax.set_xticklabels(x_labels)
    # --- LIMITI ASSE Y SPECIFICI ---
    if col == "total_SC_MWh":
        ax.set_ylim(0, 3000)

    elif col == "CSC_MWh":
        ax.set_ylim(0, 450)

    elif col == "Net_Electricity_Costs_kEUR":
        ax.set_ylim(0, 3000)

    # --- LEGENDA PER OGNI GRAFICO (SOPRA) ---
    ax.legend(
        loc="lower center",
        bbox_to_anchor=(0.5, 1.05),
        ncol=3,
        fontsize=8,
        frameon=False
    )

plt.tight_layout()

out_png = excel_path.with_name(excel_path.stem + "_PV_recap_plot_efficiency.png")
plt.savefig(out_png, dpi=400, bbox_inches="tight")
plt.show()

