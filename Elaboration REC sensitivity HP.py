from pathlib import Path
import os
import glob
import pandas as pd
import numpy as np

# =========================
# CONFIG
# =========================
name_scenario = "Retrofit_sensitivity"  # scenario CEA che contiene i risultati
hp_numbers = [4, 5, 6, 7]

# Base project path
base_path = Path(f"C:/CityEnergyAnalyst/Paper_prova/{name_scenario}/outputs/data")

# Demand folders: .../demand_HP{number}
demand_folder_template = base_path / "demand_HP{n}"

# PV folder (uguale per tutti gli HP) - MODIFICA QUI se serve
pv_folder_single = Path(r"C:\CityEnergyAnalyst\Paper_prova\Retrofit_sensitivity\outputs\data\potentials\solar_PV1")

# Se invece PV cambia per HP, usa questa e commenta pv_folder_single sopra:
# pv_folder_template = base_path / "potentials" / "solar_HP{n}"

# Output Excel
output_path = Path("C:/CityEnergyAnalyst/Paper_prova")
output_folder_community = output_path / "Elaboration_REC"
output_folder_community.mkdir(parents=True, exist_ok=True)

community_file = output_folder_community / f"community_{name_scenario}_HPcompare_RII.xlsx"


# =========================
# FUNZIONI
# =========================
def load_demand(demand_folder: Path):
    """Legge i file B*.csv e restituisce demand_final_df (Date + col per edificio) e building_ids."""
    csv_files = sorted(glob.glob(os.path.join(str(demand_folder), "B*.csv")))
    if not csv_files:
        raise FileNotFoundError(f"Nessun file B*.csv trovato in: {demand_folder}")

    dfs = []
    building_ids = []

    for i, file in enumerate(csv_files):
        df = pd.read_csv(file)

        building_id = os.path.splitext(os.path.basename(file))[0]  # es. B123
        building_ids.append(building_id)

        # Colonna di domanda (come nel tuo script originale)
        temp_df = df[["GRID_kWh"]].copy()
        temp_df.rename(columns={"GRID_kWh": building_id}, inplace=True)

        if i == 0:
            # Nel demand è 'DATE' (ma se fosse 'Date', fallback)
            date_col = "DATE" if "DATE" in df.columns else "Date"
            temp_df.insert(0, "Date", df[date_col])

        dfs.append(temp_df)

    final_df = pd.concat(dfs, axis=1)
    final_df = final_df.loc[:, ~final_df.columns.duplicated()]
    final_df["Date"] = pd.to_datetime(final_df["Date"])
    return final_df, building_ids


def load_pv(pv_folder: Path, building_ids):
    """Legge i file PV B*_PV.csv e restituisce radiation_final_df (Date + col per edificio) allineato ai building_ids."""
    csv_files = sorted(glob.glob(os.path.join(str(pv_folder), "B*_PV.csv")))
    if not csv_files:
        raise FileNotFoundError(f"Nessun file B*_PV.csv trovato in: {pv_folder}")

    dfs = []
    seen_buildings = set()

    for i, file in enumerate(csv_files):
        df = pd.read_csv(file)

        filename = os.path.splitext(os.path.basename(file))[0]  # es. B123_PV
        bld = filename.split("_")[0]  # es. B123
        seen_buildings.add(bld)

        temp_df = df[["E_PV_gen_kWh"]].copy()
        temp_df.rename(columns={"E_PV_gen_kWh": bld}, inplace=True)

        if i == 0:
            # In PV è 'Date'
            temp_df.insert(0, "Date", df["Date"])

        dfs.append(temp_df)

    final_df = pd.concat(dfs, axis=1)
    final_df = final_df.loc[:, ~final_df.columns.duplicated()]
    final_df["Date"] = pd.to_datetime(final_df["Date"]).dt.tz_localize(None)

    # Assicura stesse colonne (Date + building_ids).
    # Se mancano edifici nella PV, li aggiunge a 0.
    for b in building_ids:
        if b not in final_df.columns:
            final_df[b] = 0.0

    final_df = final_df[["Date"] + building_ids]
    return final_df


def build_time_price_df():
    """Crea df_time per 2016 (8760h), timeband e prezzi."""
    datetime_index = pd.date_range(start="2016-01-01", periods=8760, freq="h")

    df_time = pd.DataFrame({
        "Date": datetime_index,
        "Month": datetime_index.month,
        "Day": datetime_index.day,
        "Hour": datetime_index.hour,
        "Day type": datetime_index.dayofweek.map(lambda x: "Weekday" if x < 5 else ("Saturday" if x == 5 else "Sunday"))
    })

    def classify_timeband(row):
        if row["Day type"] == "Weekday":
            if 8 <= row["Hour"] < 19:
                return "F1"
            elif (7 <= row["Hour"] < 8) or (19 <= row["Hour"] < 23):
                return "F2"
            else:
                return "F3"
        elif row["Day type"] == "Saturday":
            if 7 <= row["Hour"] < 23:
                return "F2"
            else:
                return "F3"
        else:
            return "F3"

    df_time["Hourly timeband"] = df_time.apply(classify_timeband, axis=1)

    price_surplus_dict = {
        (1, "F1"): 0.10085, (2, "F1"): 0.08504, (3, "F1"): 0.08369, (4, "F1"): 0.07386, (5, "F1"): 0.08233,
        (6, "F1"): 0.09982, (7, "F1"): 0.10652, (8, "F1"): 0.11672, (9, "F1"): 0.11223, (10, "F1"): 0.11111,
        (11, "F1"): 0.13245, (12, "F1"): 0.13677,
        (1, "F2"): 0.09888, (2, "F2"): 0.08536, (3, "F2"): 0.07300, (4, "F2"): 0.07539, (5, "F2"): 0.08216,
        (6, "F2"): 0.08865, (7, "F2"): 0.10737, (8, "F2"): 0.11632, (9, "F2"): 0.10082, (10, "F2"): 0.10281,
        (11, "F2"): 0.12457, (12, "F2"): 0.12682,
        (1, "F3"): 0.08334, (2, "F3"): 0.07173, (3, "F3"): 0.05639, (4, "F3"): 0.05903, (5, "F3"): 0.06271,
        (6, "F3"): 0.07438, (7, "F3"): 0.10019, (8, "F3"): 0.11625, (9, "F3"): 0.08511, (10, "F3"): 0.09625,
        (11, "F3"): 0.10931, (12, "F3"): 0.10954,
    }
    df_time["Price surplus"] = df_time.apply(
        lambda row: price_surplus_dict.get((row["Month"], row["Hourly timeband"]), 0),
        axis=1
    )

    price_purchase_dict = {
        (1, "F1"): 0.38, (2, "F1"): 0.36, (3, "F1"): 0.34, (4, "F1"): 0.39, (5, "F1"): 0.34, (6, "F1"): 0.35,
        (7, "F1"): 0.35, (8, "F1"): 0.35, (9, "F1"): 0.36, (10, "F1"): 0.37, (11, "F1"): 0.35, (12, "F1"): 0.38,
        (1, "F2"): 0.38, (2, "F2"): 0.36, (3, "F2"): 0.34, (4, "F2"): 0.39, (5, "F2"): 0.34, (6, "F2"): 0.35,
        (7, "F2"): 0.35, (8, "F2"): 0.35, (9, "F2"): 0.36, (10, "F2"): 0.37, (11, "F2"): 0.35, (12, "F2"): 0.38,
        (1, "F3"): 0.38, (2, "F3"): 0.36, (3, "F3"): 0.34, (4, "F3"): 0.39, (5, "F3"): 0.34, (6, "F3"): 0.35,
        (7, "F3"): 0.35, (8, "F3"): 0.35, (9, "F3"): 0.36, (10, "F3"): 0.37, (11, "F3"): 0.35, (12, "F3"): 0.38,
    }
    df_time["Price purchase"] = df_time.apply(
        lambda row: price_purchase_dict.get((row["Month"], row["Hourly timeband"]), 0),
        axis=1
    )

    return df_time


def compute_hourly_flows(demand_final_df, radiation_final_df, building_ids):
    """Calcola self-consumption, import, export per edificio (DataFrame con Date + building_ids)."""
    merged_df = pd.merge(
        demand_final_df[["Date"] + building_ids],
        radiation_final_df[["Date"] + building_ids],
        on="Date",
        suffixes=("_demand", "_PV")
    )

    sc_dict = {"Date": merged_df["Date"]}
    import_dict = {"Date": merged_df["Date"]}
    export_dict = {"Date": merged_df["Date"]}

    for b in building_ids:
        demand_series = merged_df[f"{b}_demand"].to_numpy()
        pv_series = merged_df[f"{b}_PV"].to_numpy()

        self_consumption = np.where(pv_series > 0, np.minimum(demand_series, pv_series), 0.0)
        imp = np.where(demand_series > pv_series, demand_series - pv_series, 0.0)
        exp = np.where(demand_series < pv_series, pv_series - demand_series, 0.0)

        sc_dict[b] = self_consumption
        import_dict[b] = imp
        export_dict[b] = exp

    self_consumption_df = pd.DataFrame(sc_dict)
    import_df = pd.DataFrame(import_dict)
    export_df = pd.DataFrame(export_dict)
    return self_consumption_df, import_df, export_df


def compute_valutazione_CER(df_time, demand_final_df, radiation_final_df,
                            self_consumption_df, import_df, export_df, building_ids):
    """Crea valutazione_CER_df (come il tuo, ma dipendente dalla domanda per ogni HP)."""
    valutazione_CER_df = pd.DataFrame()
    valutazione_CER_df["Date"] = df_time["Date"]
    valutazione_CER_df["Month"] = df_time["Month"]
    valutazione_CER_df["Timeband"] = df_time["Hourly timeband"]
    valutazione_CER_df["Day type"] = df_time["Day type"]

    # === Queste sono le colonne che vuoi aggiornare in base allo scenario HP ===
    valutazione_CER_df["total_cons"] = demand_final_df[building_ids].sum(axis=1)
    valutazione_CER_df["total_PV"] = radiation_final_df[building_ids].sum(axis=1)
    valutazione_CER_df["total_SC"] = self_consumption_df[building_ids].sum(axis=1)

    valutazione_CER_df["SCI"] = np.where(
        valutazione_CER_df["total_PV"] > 0,
        valutazione_CER_df["total_SC"] / valutazione_CER_df["total_PV"],
        0
    )
    valutazione_CER_df["SSI"] = np.where(
        valutazione_CER_df["total_PV"] > 0,
        valutazione_CER_df["total_SC"] / valutazione_CER_df["total_cons"],
        0
    )

    valutazione_CER_df["import"] = import_df[building_ids].sum(axis=1)
    valutazione_CER_df["export"] = export_df[building_ids].sum(axis=1)
    valutazione_CER_df["CSC"] = np.minimum(valutazione_CER_df["import"], valutazione_CER_df["export"])

    valutazione_CER_df["SCI_REC"] = np.where(
        valutazione_CER_df["total_PV"] > 0,
        (valutazione_CER_df["total_SC"] + valutazione_CER_df["CSC"]) / valutazione_CER_df["total_PV"],
        0
    )
    valutazione_CER_df["SSI_REC"] = np.where(
        valutazione_CER_df["total_PV"] > 0,
        (valutazione_CER_df["total_SC"] + valutazione_CER_df["CSC"]) / valutazione_CER_df["total_cons"],
        0
    )

    # === Il resto resta uguale (prezzi / incentivi / costi / ricavi) ===
    valutazione_CER_df["Price surplus"] = df_time["Price surplus"]
    valutazione_CER_df["Price purchase"] = df_time["Price purchase"]

    valutazione_CER_df["Incentive"] = (
        (np.minimum(60 + np.maximum(0, 180 - valutazione_CER_df["Price surplus"]), 100) + 10.57) / 1000
    )

    valutazione_CER_df["Energy costs BAU"] = valutazione_CER_df["Price purchase"] * valutazione_CER_df["total_cons"]

    valutazione_CER_df["Energy costs REC"] = valutazione_CER_df["Price purchase"] * (
        valutazione_CER_df["import"] - valutazione_CER_df["CSC"]
    )

    valutazione_CER_df["Energy revenues total"] = (
        (valutazione_CER_df["Price surplus"] * (valutazione_CER_df["export"] - valutazione_CER_df["CSC"])) +
        ((valutazione_CER_df["Price surplus"] + valutazione_CER_df["Incentive"]) * valutazione_CER_df["CSC"])
    )

    valutazione_CER_df["Energy revenues REC_CSC_TIP"] = valutazione_CER_df["Incentive"] * valutazione_CER_df["CSC"]
    valutazione_CER_df["Energy revenues REC_CSC_RiD"] = valutazione_CER_df["Price surplus"] * valutazione_CER_df["CSC"]
    valutazione_CER_df["Energy revenues REC_surplus"] = (
        valutazione_CER_df["Price surplus"] * (valutazione_CER_df["export"] - valutazione_CER_df["CSC"])
    )

    return valutazione_CER_df

def summarize_for_comparison(valutazione_CER_df: pd.DataFrame, scenario_name: str) -> dict:
    """Crea una riga riassuntiva per il foglio Comparison."""
    sum_cols = [
        "total_cons", "total_PV", "total_SC", "import", "export", "CSC",
        "Energy costs BAU", "Energy costs REC",
        "Energy revenues total", "Energy revenues REC_CSC_TIP",
        "Energy revenues REC_CSC_RiD", "Energy revenues REC_surplus"
    ]

    out = {"Scenario": scenario_name}

    # somme
    for c in sum_cols:
        out[c] = float(valutazione_CER_df[c].sum())

    # media SCI solo per >0
    sci_pos = valutazione_CER_df.loc[valutazione_CER_df["SCI"] > 0, "SCI"]
    out["SCI_mean_pos"] = float(sci_pos.mean()) if len(sci_pos) > 0 else 0.0

    # media SSI (tutte le ore)
    out["SSI_mean"] = float(valutazione_CER_df["SSI"].mean())

    return out

# =========================
# RUN
# =========================
df_time = build_time_price_df()

comparison_rows = []  # <-- qui accumulo i riassunti per HP

with pd.ExcelWriter(community_file) as writer:
    for n in hp_numbers:
        demand_folder = Path(str(demand_folder_template).format(n=n))

        pv_folder = pv_folder_single

        # Load
        demand_final_df, building_ids = load_demand(demand_folder)
        radiation_final_df = load_pv(pv_folder, building_ids)

        # Allinea colonne e Date
        demand_final_df = demand_final_df[["Date"] + building_ids].copy()
        radiation_final_df = radiation_final_df[["Date"] + building_ids].copy()

        # Flussi orari
        self_consumption_df, import_df, export_df = compute_hourly_flows(
            demand_final_df, radiation_final_df, building_ids
        )

        # Valutazione CER
        valutazione_CER_df = compute_valutazione_CER(
            df_time, demand_final_df, radiation_final_df,
            self_consumption_df, import_df, export_df, building_ids
        )

        # Scrittura fogli richiesti
        demand_sheet = f"Demand_kWh_HP{n}"
        cer_sheet = f"valutazione CER_HP{n}"
        demand_final_df.to_excel(writer, sheet_name=demand_sheet, index=False)
        valutazione_CER_df.to_excel(writer, sheet_name=cer_sheet, index=False)

        # Riga per Comparison
        comparison_rows.append(
            summarize_for_comparison(valutazione_CER_df, scenario_name=f"HP{n}")
        )

    # === Foglio Comparison ===
    comparison_df = pd.DataFrame(comparison_rows)

    # (opzionale) ordina le colonne in modo leggibile
    ordered_cols = ["Scenario"] + [
        "total_cons", "total_PV", "total_SC", "import", "export", "CSC",
        "SCI_mean_pos", "SSI_mean",
        "Energy costs BAU", "Energy costs REC",
        "Energy revenues total", "Energy revenues REC_CSC_TIP",
        "Energy revenues REC_CSC_RiD", "Energy revenues REC_surplus"
    ]
    comparison_df = comparison_df[ordered_cols]

    comparison_df.to_excel(writer, sheet_name="Comparison", index=False)

print("File Excel generato:", community_file)

