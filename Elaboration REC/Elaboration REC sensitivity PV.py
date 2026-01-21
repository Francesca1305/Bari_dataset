from pathlib import Path
import os
import pandas as pd
import glob
import numpy as np

# =========================
# SETTINGS
# =========================
name_scenario = "Retrofit_I"

# Demand folder (BAU demand)
demand_folder = Path(r"C:\CityEnergyAnalyst\Paper_prova\Retrofit_I_prova\outputs\data\demand_stochastic")

# PV technologies folders: ...\potentials\solar_PV{number}
pv_numbers = [1, 5, 6, 7, 8]
pv_base_folder = Path(r"C:\CityEnergyAnalyst\Paper_prova\Retrofit_sensitivity\outputs\data\potentials")

# Output
output_folder_community = Path(r"C:\CityEnergyAnalyst\Paper_prova\Elaboration_REC\sensitivity PV")
output_folder_community.mkdir(parents=True, exist_ok=True)
community_file = output_folder_community / f"community_{name_scenario}_sensitivity_PV.xlsx"

# =========================
# 1) DEMAND (ONCE)
# =========================
csv_files_demand = glob.glob(str(demand_folder / "B*.csv"))
if not csv_files_demand:
    raise FileNotFoundError(f"Nessun file B*.csv trovato in demand_folder: {demand_folder}")

dfs_demand = []
building_ids = []

for i, file in enumerate(sorted(csv_files_demand)):
    df = pd.read_csv(file)

    building_id = os.path.splitext(os.path.basename(file))[0]
    building_ids.append(building_id)

    temp_df = df[['GRID_kWh']].copy()
    temp_df.rename(columns={'GRID_kWh': building_id}, inplace=True)

    if i == 0:
        temp_df.insert(0, 'Date', df['DATE'])

    dfs_demand.append(temp_df)

final_df_demand = pd.concat(dfs_demand, axis=1)
demand_final_df = final_df_demand.loc[:, ~final_df_demand.columns.duplicated()].copy()
demand_final_df['Date'] = pd.to_datetime(demand_final_df['Date'])

ordered_columns = ['Date'] + building_ids
demand_final_df = demand_final_df[ordered_columns]

# =========================
# 2) TIMEBANDS + PRICES (ONCE)
# =========================
datetime_index = pd.date_range(start='2016-01-01', periods=8760, freq='h')

df_time = pd.DataFrame({
    'Date': datetime_index,
    'Month': datetime_index.month,
    'Day': datetime_index.day,
    'Hour': datetime_index.hour,
    'Day type': datetime_index.dayofweek.map(
        lambda x: 'Weekday' if x < 5 else ('Saturday' if x == 5 else 'Sunday')
    )
})

def classify_timeband(row):
    if row['Day type'] == 'Weekday':
        if 8 <= row['Hour'] < 19:
            return 'F1'
        elif (7 <= row['Hour'] < 8) or (19 <= row['Hour'] < 23):
            return 'F2'
        else:
            return 'F3'
    elif row['Day type'] == 'Saturday':
        if 7 <= row['Hour'] < 23:
            return 'F2'
        else:
            return 'F3'
    else:
        return 'F3'

df_time['Hourly timeband'] = df_time.apply(classify_timeband, axis=1)

price_surplus_dict = {
    (1, 'F1'): 0.10085, (2, 'F1'): 0.08504, (3, 'F1'): 0.08369, (4, 'F1'): 0.07386, (5, 'F1'): 0.08233,
    (6, 'F1'): 0.09982, (7, 'F1'): 0.10652, (8, 'F1'): 0.11672, (9, 'F1'): 0.11223, (10, 'F1'): 0.11111,
    (11, 'F1'): 0.13245, (12, 'F1'): 0.13677,
    (1, 'F2'): 0.09888, (2, 'F2'): 0.08536, (3, 'F2'): 0.07300, (4, 'F2'): 0.07539, (5, 'F2'): 0.08216,
    (6, 'F2'): 0.08865, (7, 'F2'): 0.10737, (8, 'F2'): 0.11632, (9, 'F2'): 0.10082, (10, 'F2'): 0.10281,
    (11, 'F2'): 0.12457, (12, 'F2'): 0.12682,
    (1, 'F3'): 0.08334, (2, 'F3'): 0.07173, (3, 'F3'): 0.05639, (4, 'F3'): 0.05903, (5, 'F3'): 0.06271,
    (6, 'F3'): 0.07438, (7, 'F3'): 0.10019, (8, 'F3'): 0.11625, (9, 'F3'): 0.08511, (10, 'F3'): 0.09625,
    (11, 'F3'): 0.10931, (12, 'F3'): 0.10954,
}
df_time['Price surplus'] = df_time.apply(
    lambda row: price_surplus_dict.get((row['Month'], row['Hourly timeband']), 0), axis=1
)

price_purchase_dict = {
    (1, 'F1'): 0.38, (2, 'F1'): 0.36, (3, 'F1'): 0.34, (4, 'F1'): 0.39, (5, 'F1'): 0.34, (6, 'F1'): 0.35,
    (7, 'F1'): 0.35, (8, 'F1'): 0.35, (9, 'F1'): 0.36, (10, 'F1'): 0.37, (11, 'F1'): 0.35, (12, 'F1'): 0.38,
    (1, 'F2'): 0.38, (2, 'F2'): 0.36, (3, 'F2'): 0.34, (4, 'F2'): 0.39, (5, 'F2'): 0.34, (6, 'F2'): 0.35,
    (7, 'F2'): 0.35, (8, 'F2'): 0.35, (9, 'F2'): 0.36, (10, 'F2'): 0.37, (11, 'F2'): 0.35, (12, 'F2'): 0.38,
    (1, 'F3'): 0.38, (2, 'F3'): 0.36, (3, 'F3'): 0.34, (4, 'F3'): 0.39, (5, 'F3'): 0.34, (6, 'F3'): 0.35,
    (7, 'F3'): 0.35, (8, 'F3'): 0.35, (9, 'F3'): 0.36, (10, 'F3'): 0.37, (11, 'F3'): 0.35, (12, 'F3'): 0.38,
}
df_time['Price purchase'] = df_time.apply(
    lambda row: price_purchase_dict.get((row['Month'], row['Hourly timeband']), 0), axis=1
)

# Base valutazione (non dipende dal PV)
valutazione_base = pd.DataFrame({
    'Date': df_time['Date'],
    'Month': df_time['Month'],
    'Timeband': df_time['Hourly timeband'],
    'Day type': df_time['Day type'],
    'total_cons': demand_final_df[building_ids].sum(axis=1),
    'Price surplus': df_time['Price surplus'],
    'Price purchase': df_time['Price purchase'],
})

valutazione_base['Incentive'] = (
    (np.minimum(60 + np.maximum(0, 180 - valutazione_base['Price surplus']), 100) + 10.57) / 1000
)

# =========================
# 3) LOOP PV TECHNOLOGIES + WRITE EXCEL
# =========================
comparison_rows = []

with pd.ExcelWriter(community_file) as writer:

    for pv_n in pv_numbers:
        radiation_folder = pv_base_folder / f"solar_PV{pv_n}"
        csv_files_radiation = glob.glob(str(radiation_folder / "B*_PV.csv"))

        if not csv_files_radiation:
            print(f"ATTENZIONE: nessun file PV trovato per PV{pv_n} in {radiation_folder}. Foglio saltato.")
            continue

        # --- Read PV files for this technology ---
        pv_map = {}
        date_series = None

        for i, file in enumerate(sorted(csv_files_radiation)):
            df = pd.read_csv(file)

            filename = os.path.splitext(os.path.basename(file))[0]
            building_id = filename.split('_')[0]

            # First file provides Date
            if date_series is None:
                date_series = pd.to_datetime(df['Date']).dt.tz_localize(None)

            pv_map[building_id] = df['E_PV_gen_kWh'].values

        # Build PV dataframe with ALL buildings; missing -> 0
        radiation_final_df = pd.DataFrame({'Date': date_series})
        for b in building_ids:
            if b in pv_map:
                radiation_final_df[b] = pv_map[b]
            else:
                radiation_final_df[b] = 0.0

        radiation_final_df = radiation_final_df[ordered_columns]

        # Merge demand + PV on Date
        merged_df = pd.merge(
            demand_final_df, radiation_final_df, on='Date',
            suffixes=('_demand', '_PV')
        )

        # --- Per-building computations (vectorized via loop) ---
        sc_dict = {'Date': merged_df['Date']}
        imp_dict = {'Date': merged_df['Date']}
        exp_dict = {'Date': merged_df['Date']}

        for b in building_ids:
            demand_series = merged_df[f"{b}_demand"].to_numpy()
            pv_series = merged_df[f"{b}_PV"].to_numpy()

            sc = np.where(pv_series > 0, np.minimum(demand_series, pv_series), 0.0)
            imp = np.where(demand_series > pv_series, demand_series - pv_series, 0.0)
            exp = np.where(demand_series < pv_series, pv_series - demand_series, 0.0)

            sc_dict[b] = sc
            imp_dict[b] = imp
            exp_dict[b] = exp

        self_consumption_df = pd.DataFrame(sc_dict)
        import_df = pd.DataFrame(imp_dict)
        export_df = pd.DataFrame(exp_dict)

        # --- Valutazione CER for this PV technology ---
        valutazione_CER_df = valutazione_base.copy()

        valutazione_CER_df['total_PV'] = radiation_final_df[building_ids].sum(axis=1)
        valutazione_CER_df['total_SC'] = self_consumption_df[building_ids].sum(axis=1)

        valutazione_CER_df['SCI'] = np.where(
            valutazione_CER_df['total_PV'] > 0,
            valutazione_CER_df['total_SC'] / valutazione_CER_df['total_PV'],
            0.0
        )
        valutazione_CER_df['SSI'] = np.where(
            valutazione_CER_df['total_PV'] > 0,
            valutazione_CER_df['total_SC'] / valutazione_CER_df['total_cons'],
            0.0
        )

        valutazione_CER_df['import'] = import_df[building_ids].sum(axis=1)
        valutazione_CER_df['export'] = export_df[building_ids].sum(axis=1)
        valutazione_CER_df['CSC'] = np.minimum(valutazione_CER_df['import'], valutazione_CER_df['export'])

        valutazione_CER_df['SCI_REC'] = np.where(
            valutazione_CER_df['total_PV'] > 0,
            (valutazione_CER_df['total_SC'] + valutazione_CER_df['CSC']) / valutazione_CER_df['total_PV'],
            0.0
        )
        valutazione_CER_df['SSI_REC'] = np.where(
            valutazione_CER_df['total_PV'] > 0,
            (valutazione_CER_df['total_SC'] + valutazione_CER_df['CSC']) / valutazione_CER_df['total_cons'],
            0.0
        )

        valutazione_CER_df['Energy costs BAU'] = valutazione_CER_df['Price purchase'] * valutazione_CER_df['total_cons']
        valutazione_CER_df['Energy costs REC'] = valutazione_CER_df['Price purchase'] * (
            valutazione_CER_df['import'] - valutazione_CER_df['CSC']
        )

        valutazione_CER_df['Energy revenues total'] = (
            (valutazione_CER_df['Price surplus'] * (valutazione_CER_df['export'] - valutazione_CER_df['CSC'])) +
            ((valutazione_CER_df['Price surplus'] + valutazione_CER_df['Incentive']) * valutazione_CER_df['CSC'])
        )
        valutazione_CER_df['Energy revenues REC_CSC_TIP'] = (
            valutazione_CER_df['Incentive'] * valutazione_CER_df['CSC']
        )
        valutazione_CER_df['Energy revenues REC_CSC_RiD'] = (
            valutazione_CER_df['Price surplus'] * valutazione_CER_df['CSC']
        )
        valutazione_CER_df['Energy revenues REC_surplus'] = (
            valutazione_CER_df['Price surplus'] * (valutazione_CER_df['export'] - valutazione_CER_df['CSC'])
        )

        # Write per-technology sheet
        sheet_name = f"valutazione CER_PV{pv_n}"
        valutazione_CER_df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Build comparison row
        row = {
            "Scenario": f"PV{pv_n}",
            "total_PV_sum": float(valutazione_CER_df["total_PV"].sum()),
            "total_SC_sum": float(valutazione_CER_df["total_SC"].sum()),
            "import_sum": float(valutazione_CER_df["import"].sum()),
            "export_sum": float(valutazione_CER_df["export"].sum()),
            "CSC_sum": float(valutazione_CER_df["CSC"].sum()),
            "Energy costs BAU_sum": float(valutazione_CER_df["Energy costs BAU"].sum()),
            "Energy costs REC_sum": float(valutazione_CER_df["Energy costs REC"].sum()),
            "Energy revenues total_sum": float(valutazione_CER_df["Energy revenues total"].sum()),
            "Energy revenues REC_CSC_TIP_sum": float(valutazione_CER_df["Energy revenues REC_CSC_TIP"].sum()),
            "Energy revenues REC_CSC_RiD_sum": float(valutazione_CER_df["Energy revenues REC_CSC_RiD"].sum()),
            "Energy revenues REC_surplus_sum": float(valutazione_CER_df["Energy revenues REC_surplus"].sum()),
        }

        sci_pos = valutazione_CER_df.loc[valutazione_CER_df["SCI"] > 0, "SCI"]
        row["SCI_mean_pos"] = float(sci_pos.mean()) if not sci_pos.empty else 0.0

        row["SSI_mean"] = float(valutazione_CER_df["SSI"].mean())

        comparison_rows.append(row)

    # Write Comparison sheet once
    comparison_df = pd.DataFrame(comparison_rows)

    if not comparison_df.empty:
        comparison_df = comparison_df[
            [
                "Scenario",
                "total_PV_sum", "total_SC_sum", "import_sum", "export_sum", "CSC_sum",
                "SCI_mean_pos", "SSI_mean",
                "Energy costs BAU_sum", "Energy costs REC_sum",
                "Energy revenues total_sum",
                "Energy revenues REC_CSC_TIP_sum", "Energy revenues REC_CSC_RiD_sum", "Energy revenues REC_surplus_sum",
            ]
        ]
        comparison_df.to_excel(writer, sheet_name="Comparison", index=False)
    else:
        # still create an empty Comparison sheet with headers
        empty_cols = [
            "Scenario",
            "total_PV_sum", "total_SC_sum", "import_sum", "export_sum", "CSC_sum",
            "SCI_mean_pos", "SSI_mean",
            "Energy costs BAU_sum", "Energy costs REC_sum",
            "Energy revenues total_sum",
            "Energy revenues REC_CSC_TIP_sum", "Energy revenues REC_CSC_RiD_sum", "Energy revenues REC_surplus_sum",
        ]
        pd.DataFrame(columns=empty_cols).to_excel(writer, sheet_name="Comparison", index=False)

print("File Excel generato:", community_file)

