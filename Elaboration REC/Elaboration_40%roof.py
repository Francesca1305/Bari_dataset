from pathlib import Path
import os
import pandas as pd
import glob
from datetime import datetime, timedelta
import numpy as np
import json

name_scenario = "BAU_40%roof"

# Base project path
#base_path = Path(f"C:/CityEnergyAnalyst/Paper_prova/{name_scenario}/outputs/data")
demand_folder = r"D:\PhD\Simulazioni CEA Articolo CEES\stochastic BAU\demand_stochastic"
radiation_folder = r"D:\PhD\Simulazioni CEA Articolo CEES\Retrofit_sensitivity\outputs\data\solar-radiation"

# Define specific folders and files for community aggregated files
output_path = Path(f"{name_scenario}")
output_folder_community = (output_path / "VES")

# Create output directory if it doesn't exist
output_folder_community.mkdir(parents=True, exist_ok=True)

# Output file paths
community_file = (output_folder_community /
                  f"community_bybuilding_{name_scenario}.xlsx")

### Demand ###
csv_files_demand = glob.glob(os.path.join(demand_folder, 'B*.csv'))
dfs_demand = []
building_ids = []

# Variabile per salvare la colonna DATE (la prendiamo dal primo file)
date_column = None

# Leggiamo tutti i file di demand
for i, file in enumerate(csv_files_demand):
    df = pd.read_csv(file)

    # Estrai il nome dell'edificio dal nome del file
    building_id = os.path.splitext(os.path.basename(file))[0]
    building_ids.append(building_id)

    # Prendi solo la colonna E_sys_kWh e rinominala con il building_id
    temp_df = df[['GRID_kWh']].copy()
    temp_df.rename(columns={'GRID_kWh': building_id}, inplace=True)

    # Alla prima iterazione prendiamo anche la colonna DATE
    if i == 0:
        temp_df.insert(0, 'Date', df['DATE'])

    dfs_demand.append(temp_df)

# Ora uniamo tutti i DataFrame sulle colonne
final_df_demand = pd.concat(dfs_demand, axis=1)
# Rimuoviamo eventuali colonne duplicate di DATE (per sicurezza)
demand_final_df = final_df_demand.loc[:, ~final_df_demand.columns.duplicated()]
demand_final_df['Date'] = pd.to_datetime(demand_final_df['Date'])


### PV Generation ###
csv_files_radiation = glob.glob(os.path.join(radiation_folder, 'B*_radiation.csv'))
dfs_radiation = []
rad_m2_dict = {}  # per salvare radiazione al m2 per edificio

for i, file in enumerate(csv_files_radiation):
    df = pd.read_csv(file)

    filename = os.path.splitext(os.path.basename(file))[0]  # es. B123_radiation
    building_id = filename.split('_')[0]  # es. B123

    # Calcolo radiazione oraria al m2
    rad_m2 = df['roofs_top_kW'] / df['roofs_top_m2']
    rad_m2_dict[building_id] = rad_m2.values

    # Calcolo produzione PV oraria:
    # rad_m2 * (40% di roofs_top_m2) * efficienza(0.20) * performance_ratio(0.8)
    pv_area = df['roofs_top_m2'] * 0.40
    pv_gen = rad_m2 * pv_area * 0.20 * 0.8

    temp_df = pv_gen.to_frame(name=building_id)

    if i == 0:
        temp_df.insert(0, 'Date', df['Date'])

    dfs_radiation.append(temp_df)

# Salva rad_m2 in Excel
rad_m2_df = pd.DataFrame(rad_m2_dict)
rad_m2_output = output_folder_community / "rad_m2_building.xlsx"
rad_m2_df.to_excel(rad_m2_output, index=False)
print("File radiazione al m2 generato:", rad_m2_output)

# ── RECAP PER EDIFICIO ──────────────────────────────────────────────
pv_recap_rows = []

for file in csv_files_radiation:
    df = pd.read_csv(file)
    filename = os.path.splitext(os.path.basename(file))[0]
    building_id = filename.split('_')[0]

    rad_m2      = df['roofs_top_kW'] / df['roofs_top_m2']
    pv_area     = df['roofs_top_m2'] * 0.40          # m2 installati (costante per riga)
    pv_gen      = rad_m2 * pv_area * 0.20 * 0.8

    pv_recap_rows.append({
        'Building': building_id,
        'Rad_total_annual_kWh': df['roofs_top_kW'].sum(), # radiazione totale annua [kWh]
        'PV_area_m2': pv_area.iloc[0],                    # area PV installata [m2]
        'n_panels': pv_area.iloc[0] / 1.76,
        'kW_installed': (pv_area.iloc[0] / 1.76)*0.3257,  # potenza installata [kWp]
        'PV_gen_annual_kWh': pv_gen.sum(),                # produzione PV annua [kWh]
    })

pv_recap_df = pd.DataFrame(pv_recap_rows).set_index('Building')

pv_recap_output = output_folder_community / "PV_buildings.xlsx"
pv_recap_df.to_excel(pv_recap_output)
print("File recap PV generato:", pv_recap_output)


# Costruisci radiation_final_df con la produzione PV calcolata
final_df_radiation = pd.concat(dfs_radiation, axis=1)
radiation_final_df = final_df_radiation.loc[:, ~final_df_radiation.columns.duplicated()]
radiation_final_df['Date'] = pd.to_datetime(radiation_final_df['Date']).dt.tz_localize(None)

### Ordiniamo le Colonne ###
ordered_columns = ['Date'] + building_ids
demand_final_df = demand_final_df[ordered_columns]
radiation_final_df = radiation_final_df[ordered_columns]

# Facciamo un merge per essere sicuri che siano allineati sulle date
merged_df = pd.merge(demand_final_df, radiation_final_df, on='Date', suffixes=('_demand', '_PV'))

# Creiamo un dizionario per raccogliere tutte le serie
self_consumption_dict = {'Date': merged_df['Date']}
SCI_dict = {'Date': merged_df['Date']}
SSI_dict = {'Date': merged_df['Date']}
import_kWh_dict = {'Date': merged_df['Date']}
export_kWh_dict = {'Date': merged_df['Date']}

# Per ogni edificio
for col in demand_final_df.columns:
    if col != 'Date':  # saltiamo la colonna Date
        demand_series = merged_df[f"{col}_demand"]
        pv_series = merged_df[f"{col}_PV"]

        # Calcolo autoconsumo orario
        self_consumption = np.where(pv_series > 0, np.minimum(demand_series, pv_series), 0)
        SCI = np.where(pv_series > 0, self_consumption/pv_series, 0)
        SSI = np.where(pv_series > 0, self_consumption/demand_series, 0)
        import_kWh = np.where(demand_series > pv_series, demand_series - pv_series, 0)
        export_kWh = np.where(demand_series < pv_series, pv_series - demand_series, 0)

        # Aggiungiamo al dizionario
        self_consumption_dict[col] = self_consumption
        SCI_dict[col] = SCI
        SSI_dict[col] = SSI
        import_kWh_dict[col] = import_kWh
        export_kWh_dict[col] = export_kWh

# Ora creiamo il dataframe tutto in una volta (veloce e senza frammentazione)
self_consumption_df = pd.DataFrame(self_consumption_dict)
SCI_df = pd.DataFrame(SCI_dict)
SSI_df = pd.DataFrame(SSI_dict)
import_df = pd.DataFrame(import_kWh_dict)
export_df = pd.DataFrame(export_kWh_dict)


### Valutazione REC ###
datetime_index = pd.date_range(start='2016-01-01', periods=8760, freq='h')

df_time = pd.DataFrame({
    'Date': datetime_index,
    'Month': datetime_index.month,
    'Day': datetime_index.day,
    'Hour': datetime_index.hour,
    'Day type': datetime_index.dayofweek.map(lambda x: 'Weekday' if x < 5 else ('Saturday' if x == 5 else 'Sunday'))
})

# Definire le fasce orarie
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
    else:  # Sunday
        return 'F3'

df_time['Hourly timeband'] = df_time.apply(classify_timeband, axis=1)

# Prezzi surplus
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

df_time['Price surplus'] = df_time.apply(lambda row: price_surplus_dict.get((row['Month'], row['Hourly timeband']), 0), axis=1)

# Prezzi acquisto
# price_purchase_dict = {
#     (1, 'F1'): 0.11, (2, 'F1'): 0.10, (3, 'F1'): 0.09, (4, 'F1'): 0.09, (5, 'F1'): 0.09, (6, 'F1'): 0.10,
#     (7, 'F1'): 0.11, (8, 'F1'): 0.12, (9, 'F1'): 0.12, (10, 'F1'): 0.12, (11, 'F1'): 0.15, (12, 'F1'): 0.16,
#     (1, 'F2'): 0.11, (2, 'F2'): 0.09, (3, 'F2'): 0.09, (4, 'F2'): 0.10, (5, 'F2'): 0.11, (6, 'F2'): 0.12,
#     (7, 'F2'): 0.13, (8, 'F2'): 0.15, (9, 'F2'): 0.13, (10, 'F2'): 0.13, (11, 'F2'): 0.14, (12, 'F2'): 0.15,
#     (1, 'F3'): 0.09, (2, 'F3'): 0.08, (3, 'F3'): 0.08, (4, 'F3'): 0.08, (5, 'F3'): 0.09, (6, 'F3'): 0.10,
#     (7, 'F3'): 0.10, (8, 'F3'): 0.12, (9, 'F3'): 0.11, (10, 'F3'): 0.11, (11, 'F3'): 0.12, (12, 'F3'): 0.12,
# }

price_purchase_dict = {
    (1, 'F1'): 0.38, (2, 'F1'): 0.36, (3, 'F1'): 0.34, (4, 'F1'): 0.39, (5, 'F1'): 0.34, (6, 'F1'): 0.35,
    (7, 'F1'): 0.35, (8, 'F1'): 0.35, (9, 'F1'): 0.36, (10, 'F1'): 0.37, (11, 'F1'): 0.35, (12, 'F1'): 0.38,
    (1, 'F2'): 0.38, (2, 'F2'): 0.36, (3, 'F2'): 0.34, (4, 'F2'): 0.39, (5, 'F2'): 0.34, (6, 'F2'): 0.35,
    (7, 'F2'): 0.35, (8, 'F2'): 0.35, (9, 'F2'): 0.36, (10, 'F2'): 0.37, (11, 'F2'): 0.35, (12, 'F2'): 0.38,
    (1, 'F3'): 0.38, (2, 'F3'): 0.36, (3, 'F3'): 0.34, (4, 'F3'): 0.39, (5, 'F3'): 0.34, (6, 'F3'): 0.35,
    (7, 'F3'): 0.35, (8, 'F3'): 0.35, (9, 'F3'): 0.36, (10, 'F3'): 0.37, (11, 'F3'): 0.35, (12, 'F3'): 0.38,
}

df_time['Price purchase'] = df_time.apply(lambda row: price_purchase_dict.get((row['Month'], row['Hourly timeband']), 0), axis=1)

### Prices_purchase (versione ottimizzata) ###
# Crea un DataFrame in un unico passaggio
price_purchase_values = df_time['Price purchase'].values
price_surplus_values = df_time['Price surplus'].values

# Crea un dizionario: ogni edificio → array dei prezzi
price_purchase_dict = {b: price_purchase_values for b in building_ids}
price_surplus_dict = {b: price_surplus_values for b in building_ids}

# Converte tutto in DataFrame, aggiungendo la colonna 'Date'
prices_purchase_df = pd.DataFrame(price_purchase_dict)
prices_purchase_df.insert(0, 'Date', df_time['Date'].values)
prices_surplus_df = pd.DataFrame(price_surplus_dict)
prices_surplus_df.insert(0, 'Date', df_time['Date'].values)

# Calcolo vettoriale: moltiplicazione tra array orario dei prezzi e dataframe delle domande
price_array = df_time['Price purchase'].values.reshape(-1, 1)
energy_costs_values = demand_final_df[building_ids].values * price_array

# Creazione del DataFrame finale in un solo passaggio
energy_costs_df = pd.DataFrame(energy_costs_values, columns=building_ids)
energy_costs_df.insert(0, 'Date', df_time['Date'].values)


# Ora creiamo il dataframe di valutazione CER
valutazione_CER_df = pd.DataFrame()
valutazione_CER_df['Date'] = df_time['Date']
valutazione_CER_df['Month'] = df_time['Month']
valutazione_CER_df['Timeband'] = df_time['Hourly timeband']
valutazione_CER_df['Day type'] = df_time['Day type']
valutazione_CER_df['total_cons'] = demand_final_df[building_ids].sum(axis=1)
valutazione_CER_df['total_PV'] = radiation_final_df[building_ids].sum(axis=1)
valutazione_CER_df['total_SC'] = self_consumption_df[building_ids].sum(axis=1)
valutazione_CER_df['SCI'] = np.where(valutazione_CER_df['total_PV'] > 0, valutazione_CER_df['total_SC']/
                                    valutazione_CER_df['total_PV'], 0)
valutazione_CER_df['SSI'] = np.where(valutazione_CER_df['total_PV'] > 0, valutazione_CER_df['total_SC'] /
                                     valutazione_CER_df['total_cons'], 0)
valutazione_CER_df['import'] = import_df[building_ids].sum(axis=1)
valutazione_CER_df['export'] = export_df[building_ids].sum(axis=1)
valutazione_CER_df['CSC'] = np.minimum(valutazione_CER_df['import'], valutazione_CER_df['export'])
valutazione_CER_df['SCI_REC'] = np.where(valutazione_CER_df['total_PV'] > 0, (valutazione_CER_df['total_SC']+
                                                                              valutazione_CER_df['CSC'])/
                                    valutazione_CER_df['total_PV'], 0)
valutazione_CER_df['SSI_REC'] = np.where(valutazione_CER_df['total_PV'] > 0, (valutazione_CER_df['total_SC']+
                                                                              valutazione_CER_df['CSC']) /
                                     valutazione_CER_df['total_cons'], 0)
valutazione_CER_df['Price surplus'] = df_time['Price surplus']
valutazione_CER_df['Price purchase'] = df_time['Price purchase']
valutazione_CER_df['Incentive'] = ((np.minimum(60 + np.maximum(0, 180 - valutazione_CER_df['Price surplus']),100)
                                    +10.57)/1000)
valutazione_CER_df['Energy costs BAU'] = valutazione_CER_df['Price purchase'] * valutazione_CER_df['total_cons']

valutazione_CER_df['Energy costs REC'] = valutazione_CER_df['Price purchase'] * (
    valutazione_CER_df['import'] - valutazione_CER_df['CSC']
)
valutazione_CER_df['Energy revenues total'] = ((valutazione_CER_df['Price surplus'] * (valutazione_CER_df['export'] -
                                                                                    valutazione_CER_df['CSC'])) +
                                             ((valutazione_CER_df['Price surplus']+valutazione_CER_df['Incentive']) *
                                              (valutazione_CER_df['CSC'])))
valutazione_CER_df['Energy revenues REC_CSC_TIP'] = ((valutazione_CER_df['Incentive']) *
                                             (valutazione_CER_df['CSC']))
valutazione_CER_df['Energy revenues REC_CSC_RiD'] = ((valutazione_CER_df['Price surplus']) *
                                             (valutazione_CER_df['CSC']))
valutazione_CER_df['Energy revenues REC_surplus'] = (valutazione_CER_df['Price surplus'] *
                                                     (valutazione_CER_df['export'] - valutazione_CER_df['CSC']))

# Scriviamo in Excel
with pd.ExcelWriter(community_file) as writer:
    demand_final_df.to_excel(writer, sheet_name="Demand_kWh", index=False)
    radiation_final_df.to_excel(writer, sheet_name="PV_kWh", index=False)
    self_consumption_df.to_excel(writer, sheet_name="Self_consumption_kWh", index=False)
    import_df.to_excel(writer, sheet_name="Import_kWh", index=False)
    export_df.to_excel(writer, sheet_name="Export_kWh", index=False)
    prices_purchase_df.to_excel(writer, sheet_name="Prices_purchase", index=False)
    prices_surplus_df.to_excel(writer, sheet_name="Prices_surplus", index=False)
    SCI_df.to_excel(writer, sheet_name="SCI", index=False)
    SSI_df.to_excel(writer, sheet_name="SSI", index=False)
    energy_costs_df.to_excel(writer, sheet_name="Energy Costs", index=False)
    valutazione_CER_df.to_excel(writer, sheet_name="valutazione CER", index=False)

print("File Excel generato:", community_file)
