import os
import json
import pandas as pd
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import numpy as np

excel_file = r"C:\Users\franc\Politecnico Di Torino Studenti Dropbox\Francesca Vecchi\PhD\Progetto di ricerca PhD\PhD thesis\Bari\Profiles\Probabilistic profiles.xlsx"
output_dir = "output chart profiles"
os.makedirs(output_dir, exist_ok=True)

sheet_names = ["DHW", "Occupancy", "Appliances", "Lighting"]

profiles = [
    "Single worker",
    "Single retired",
    "Working couple",
    "Retired couple",
    "Family with 3 members",
    "More than 3 members"
]

colors = {
    "Single worker": "deepskyblue",
    "Single retired": "darkorange",
    "Working couple": "gray",
    "Retired couple": "gold",
    "Family with 3 members": "blue",
    "More than 3 members": "green"
}

for sheet in sheet_names:
    df = pd.read_excel(excel_file, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    df["day_type"] = df["day_type"].astype(str).str.strip()

    n = len(df)
    x = np.arange(n)

    fig, ax = plt.subplots(figsize=(14, 7))

    for profile in profiles:
        if profile in df.columns:
            ax.plot(
                x,
                df[profile],
                label=profile,
                color=colors[profile],
                linewidth=2.0,
                solid_capstyle="round",
                solid_joinstyle="round",
                linestyle="dashed"
            )

    ax.set_ylim(0, 1)
    ax.set_title("", fontsize=11)
    ax.set_xlabel("", fontsize=11)
    ax.set_ylabel("Hourly probability", fontsize=11)
    ax.grid(False)

    hour_ticks = np.arange(0, n, 2)
    hour_labels = [str(i % 24) for i in hour_ticks]
    ax.set_xticks(hour_ticks)
    ax.set_xticklabels(hour_labels)

    groups = []
    start = 0
    current = df.loc[0, "day_type"]

    for i in range(1, n):
        if df.loc[i, "day_type"] != current:
            groups.append((current, start, i - 1))
            start = i
            current = df.loc[i, "day_type"]
    groups.append((current, start, n - 1))

    secax = ax.secondary_xaxis("bottom")
    centers = [(g[1] + g[2]) / 2 for g in groups]
    labels = [g[0] for g in groups]
    secax.set_xticks(centers)
    secax.set_xticklabels(labels)

    secax.spines["bottom"].set_visible(False)
    secax.tick_params(axis="x", length=0, pad=22)

    for _, start, end in groups[:-1]:
        ax.axvline(
            end + 0.5,
            color="grey",
            linewidth=0.7,
            linestyle= (0, (5, 10))
        )

    for spine in ["top", "right", "bottom", "left"]:
        ax.spines[spine].set_visible(True)
        ax.spines[spine].set_linewidth(1.0)

    ax.legend(
        loc="lower center",
        bbox_to_anchor=(0.5, 1.02),
        ncol=6,
        fontsize=11,
        frameon=False
    )

    plt.tight_layout()
    output_file = os.path.join(output_dir, f"{sheet}.jpeg")
    plt.savefig(output_file, format="jpeg", dpi=300, bbox_inches="tight")
    plt.close()

print("Grafici salvati nella cartella output/")