from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# --- INPUT ---
excel_path = Path(
    r"C:\CityEnergyAnalyst\Paper_prova\Elaboration_REC\sensitivity HP\community_Retrofit_sensitivity_HPcompare_RII.xlsx"
)
sheet_name = "Comparison"

# --- LOAD ---
df = pd.read_excel(excel_path, sheet_name=sheet_name)

# Ordina per HP
df["HP_num"] = df["Scenario"].astype(str).str.extract(r"HP(\d+)").astype(int)
df = df.sort_values("HP_num").reset_index(drop=True)

# Mappa COP
cop_map = {4: 3.0, 5: 3.5, 6: 4.0, 7: 4.5}
df["COP"] = df["HP_num"].map(cop_map)

# Conversioni unità
df["total_cons_MWh"] = df["total_cons"] / 1000
df["total_SC_MWh"] = df["total_SC"] / 1000
df["CSC_MWh"] = df["CSC"] / 1000

# Net Energy Costs [k€/y]
df["Net Energy Costs_kEUR"] = (df["Energy costs REC"] - df["Energy revenues total"]) / 1000

x = df["COP"].values

# --- PLOT ---
fig, axes = plt.subplots(2, 2, figsize=(8, 8))
axes = axes.flatten()  # per iterare come lista


plots = [
    ("total_cons_MWh", "Electricity consumption", "MWh/y", "lightgreen"),
    ("total_SC_MWh", "Physical self-consumption", "MWh/y", "orange"),
    ("CSC_MWh", "Collective self-consumption (CSC)", "MWh/y", "gold"),
    ("Net Energy Costs_kEUR", "REC Net Electricity Costs (costs − revenues)", "k€/y", "blue"),
]

for ax, (col, title, unit, color) in zip(axes, plots):

    if color is None:
        ax.plot(x, df[col].values, marker="o")
    else:
        ax.plot(x, df[col].values, marker="o", color=color)

    ax.set_title(title, fontsize = 10)
    ax.set_xlabel("COP value", fontsize = 10)
    ax.tick_params(axis="both", labelsize=8)
    ax.set_ylabel(unit, fontsize = 8)

    # Griglia principale בלבד
    ax.grid(True, which="major", linestyle="-", linewidth=0.6, alpha=0.6)

    # Tick x solo sui COP simulati
    ax.set_xticks(x)

    # --- Y ticks personalizzati ---
    if col == "total_cons_MWh":
        ax.set_yticks(np.arange(8800, df[col].max() + 200, 200))

    elif col == "total_SC_MWh":
        ax.set_yticks(np.arange(2225, df[col].max() + 5, 5))

    elif col == "CSC_MWh":
        ax.set_yticks(np.arange(303, df[col].max() + 1, 1))

    elif col == "Net Energy Costs_kEUR":
        y_min = np.floor(df[col].min() / 100) * 100
        y_max = np.ceil(df[col].max() / 100) * 100
        ax.set_yticks(np.arange(y_min, y_max + 100, 100))


plt.tight_layout()

# --- SAVE ---
out_png = excel_path.with_name(excel_path.stem + "_recap_plot.png")
plt.savefig(out_png, dpi=300, bbox_inches="tight")
plt.show()

print("Grafico salvato in:", out_png)