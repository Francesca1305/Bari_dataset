from pathlib import Path
import os
import pandas as pd
import glob
from datetime import datetime, timedelta
import numpy as np
import json

name_scenario = "Retrofit_II_prova"

# Base project path
base_path = Path(f"C:/CityEnergyAnalyst/Paper_prova/{name_scenario}/outputs/data")
demand_folder = (base_path / "demand_stochastic")
radiation_folder = "C:/CityEnergyAnalyst/Paper_prova/BAU_scenario_prova3/outputs/data/potentials/solar"

# Define specific folders and files for community aggregated files
output_path = Path("C:/CityEnergyAnalyst/Paper_prova")
output_folder_community = (output_path / "Elaboration_NG")

# Create output directory if it doesn't exist
output_folder_community.mkdir(parents=True, exist_ok=True)

# Output file paths
community_file = (output_folder_community /
                  f"community_bybuilding_{name_scenario}_NG_stochastic.xlsx")

### Demand heating ###
csv_files_demand = glob.glob(os.path.join(demand_folder, 'B*.csv'))
dfs_demand_NG_heating = []
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
    temp_df = df[['NG_hs_kWh']].copy()
    temp_df.rename(columns={'NG_hs_kWh': building_id}, inplace=True)

    # Alla prima iterazione prendiamo anche la colonna DATE
    if i == 0:
        temp_df.insert(0, 'Date', df['DATE'])

    dfs_demand_NG_heating.append(temp_df)

# Ora uniamo tutti i DataFrame sulle colonne
final_df_demand_heating = pd.concat(dfs_demand_NG_heating, axis=1)
# Rimuoviamo eventuali colonne duplicate di DATE (per sicurezza)
demand_heating_final_df = final_df_demand_heating.loc[:, ~final_df_demand_heating.columns.duplicated()]
demand_heating_final_df['Date'] = pd.to_datetime(demand_heating_final_df['Date'])


dfs_demand_NG_DHW = []
building_ids = []

# Domestic Hot Water
date_column = None

# Leggiamo tutti i file di demand
for i, file in enumerate(csv_files_demand):
    df = pd.read_csv(file)

    # Estrai il nome dell'edificio dal nome del file
    building_id = os.path.splitext(os.path.basename(file))[0]
    building_ids.append(building_id)

    # Prendi solo la colonna E_sys_kWh e rinominala con il building_id
    temp_df = df[['NG_ww_kWh']].copy()
    temp_df.rename(columns={'NG_ww_kWh': building_id}, inplace=True)

    # Alla prima iterazione prendiamo anche la colonna DATE
    if i == 0:
        temp_df.insert(0, 'Date', df['DATE'])

    dfs_demand_NG_DHW.append(temp_df)

# Ora uniamo tutti i DataFrame sulle colonne
final_df_demand_DHW = pd.concat(dfs_demand_NG_DHW, axis=1)
# Rimuoviamo eventuali colonne duplicate di DATE (per sicurezza)
demand_DHW_final_df = final_df_demand_DHW.loc[:, ~final_df_demand_DHW.columns.duplicated()]
demand_DHW_final_df['Date'] = pd.to_datetime(demand_DHW_final_df['Date'])


### Ordiniamo le Colonne ###
ordered_columns = ['Date'] + building_ids
demand_heating_final_df = demand_heating_final_df[ordered_columns]
demand_DHW_final_df = demand_DHW_final_df[ordered_columns]

# Facciamo un merge per essere sicuri che siano allineati sulle date
merged_df = pd.merge(demand_heating_final_df, demand_DHW_final_df, on='Date', suffixes=('_demand', '_demand'))


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


# Scriviamo in Excel
with pd.ExcelWriter(community_file) as writer:
    demand_heating_final_df.to_excel(writer, sheet_name="Heating_NG", index=False)
    demand_DHW_final_df.to_excel(writer, sheet_name="DHW_NG", index=False)


print("File Excel generato:", community_file)