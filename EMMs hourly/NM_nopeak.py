import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path
import json

# Upload the excel file which has the hourly values divided by building
file_path = Path(r"C:\Users\franc\PythonProject\Bari_dataset\Elaboration REC\BAU_40%roof\VES\community_bybuilding_BAU_40%roof.xlsx")
output_file_NM_nopeak = Path(r"C:\Users\franc\Desktop\ABM Bari\Elaboration_REC\NM_nopeak.xlsx")

# === PATH ===
file_json = Path(r"C:\Users\franc\PythonProject\Agent-Based-Model\Bari_elaboration\Buildings_Bari\Buildings_data_Bari.json")
output_file_costs = Path(r"C:\Users\franc\Desktop\ABM Bari\Elaboration_REC\Building_costs.xlsx")

# === LOAD JSON ===
with open(file_json, 'r') as f:
    building_data = json.load(f)

# === BUILDING IDS ===
building_ids = list(building_data.keys())

# === IDENTIFY RESIDENTIAL ===
is_residential = {b_id: (data["building_type"] == "1000")
    for b_id, data in building_data.items()}

# === LOAD COMMUNITY PRICES ===
df_rec = pd.read_excel(file_path, sheet_name="valutazione CER")

# Assumo che la colonna si chiami "Prices"
community_prices = df_rec["Price purchase"].values

# === CREATE DATETIME INDEX (2016) ===
start_date = datetime(2016, 1, 1)
dates = [start_date + timedelta(hours=i) for i in range(8760)]

# === CREATE PRICES DF ===
df_prices = pd.DataFrame(index=range(8760))
df_prices["Date"] = dates

for b in building_ids:
    col_name = f"B{b}"
    if is_residential[b]:
        df_prices[col_name] = community_prices
    else:
        df_prices[col_name] = 0.062

# === CREATE SYSTEM ACCESS DF ===
df_sys = pd.DataFrame(index=range(8760))
df_sys["Date"] = dates

for b in building_ids:
    col_name = f"B{b}"
    df_sys[col_name] = 0.0

# === SAVE EXCEL ===
with pd.ExcelWriter(output_file_costs, engine='openpyxl') as writer:
    df_prices.to_excel(writer, sheet_name="Prices", index=False)
    df_sys.to_excel(writer, sheet_name="System Access Charges", index=False)


# # # NET METERING WITH 50 kW peak threshold
# # Upload data in Excel
df_import = pd.read_excel(file_path, sheet_name="Import_kWh")
df_export = pd.read_excel(file_path, sheet_name="Export_kWh")
df_consumption = pd.read_excel(file_path, sheet_name="Demand_kWh")
df_production = pd.read_excel(file_path, sheet_name="PV_kWh")
df_selfconsumption = pd.read_excel(file_path, sheet_name="Self_consumption_kWh")
df_initial_costs = pd.read_excel(file_path, sheet_name="Import_costs")
df_syst_charges = pd.read_excel(output_file_costs, sheet_name="System Access Charges")
df_prices = pd.read_excel(output_file_costs, sheet_name="Prices")

df_consumption['Date'] = pd.to_datetime(df_consumption['Date'])
df_production['Date'] = pd.to_datetime(df_production['Date'])
df_selfconsumption['Date'] = pd.to_datetime(df_selfconsumption['Date'])
df_import['Date'] = pd.to_datetime(df_import['Date'])
df_export['Date'] = pd.to_datetime(df_export['Date'])
df_prices['Date'] = pd.to_datetime(df_prices['Date'])
df_syst_charges['Date'] = pd.to_datetime(df_syst_charges['Date'])
df_initial_costs['Date'] = pd.to_datetime(df_initial_costs['Date'])

# 🔥 SET DATE AS INDEX
df_consumption = df_consumption.set_index('Date')
df_production = df_production.set_index('Date')
df_selfconsumption = df_selfconsumption.set_index('Date')
df_import = df_import.set_index('Date')
df_export = df_export.set_index('Date')
df_prices = df_prices.set_index('Date')
df_syst_charges = df_syst_charges.set_index('Date')
df_initial_costs = df_initial_costs.set_index('Date')
building_col_ids = [f"B{b}" for b in building_ids]

# Verifica che tutti gli edifici siano presenti in Prices
price_cols = df_prices.columns.tolist()
missing = [c for c in building_col_ids if c not in price_cols]
if missing:
    raise ValueError(f"Edifici mancanti nel foglio Prices: {missing}")

# ✅ Moltiplica DataFrame × DataFrame: pandas allinea automaticamente per nome colonna
# Entrambi hanno Date come index → nessun problema di shape
initial_BAU_energy_costs = df_consumption * df_prices

# Function to calculate monthly summation of import and export values
def calculate_monthly_sum(df):
    df_monthly = df.groupby(df.index.month).sum()  # Somma per mese
    return df_monthly

import_monthly = calculate_monthly_sum(df_import)
export_monthly = calculate_monthly_sum(df_export)
production_monthly = calculate_monthly_sum(df_production)
PSC_monthly_kWh = calculate_monthly_sum(df_selfconsumption)
monthly_prices = df_prices.groupby(df_prices.index.month).mean()
system_access_monthly = df_syst_charges.groupby(df_syst_charges.index.month).sum()
consumption_monthly = calculate_monthly_sum(df_consumption)
initial_costs_after_PSC = calculate_monthly_sum(df_initial_costs)
initial_BAU_energy_costs = calculate_monthly_sum(initial_BAU_energy_costs)

# # NET METERING NO 50 kW peak threshold
def calculate_credit_distribution_nopeak(import_monthly, export_monthly,
                                         monthly_average_price,
                                         monthly_system_access,
                                         production_monthly, PSC_monthly_kWh):
    credit_dist = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)
    unused_credits = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)
    balance_kWh = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)
    balance_costs = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)
    deficit_costs_monthly = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)
    deficit_monthly_total = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)

    self_consumed_energy = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)
    energy_shared = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)

    # ✅ monthly energy costs actually paid (deficit after credits) + system access
    energy_costs_from_deficit = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)

    # ✅ self_consumed_total = PV production - unused credits (end of month)
    pv_prod = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)
    PSC_kWh = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)
    monthly_access_charges = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)
    self_consumed_total = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)
    self_consumed_onlyNM = pd.DataFrame(0.0, index=import_monthly.index, columns=import_monthly.columns, dtype=float)

    for month in range(1, 13):
        for building in import_monthly.columns:
            import_val = import_monthly.loc[month, building]
            export_val = export_monthly.loc[month, building]

            #available_credits = unused_credits.loc[month - 1, building] if month > 1 else 0.0
            export_from_previous_month = export_monthly.loc[month - 1, building] if month > 1 else 0.0
            price = monthly_average_price.loc[month, building]

            # tuoi balance
            balance_kWh_val = import_val - export_from_previous_month
            #balance_kWh_val = import_val - (export_from_previous_month + available_credits)

            balance_kWh.loc[month, building] = balance_kWh_val
            balance_costs.loc[month, building] = balance_kWh_val * price

            # energia totale disponibile come crediti questo mese
            # total_credits_available = export_val + available_credits

            # self-consumed = crediti effettivamente usati per coprire import
            # used_for_self = min(import_val, total_credits_available)
            used_for_self = min(import_val, export_from_previous_month)

            # energy shared = crediti che restano (vanno in unused_credits)
            #shared_energy_val = max(total_credits_available - import_val, 0.0)
            shared_energy_val = max(export_val-used_for_self, 0.0)

            # deficit = import non coperto da export + crediti
            # if import_val > total_credits_available:
            #     deficit = import_val - total_credits_available
            #     credit_dist.loc[month, building] = 0.0
            #     unused_credits.loc[month, building] = 0.0
            # else:
            #     deficit = 0.0
            #     credit_dist.loc[month, building] = float(total_credits_available - import_val)
            #     unused_credits.loc[month, building] = float(credit_dist.loc[month, building])
            #
            # deficit_monthly_total.loc[month, building] = deficit
            deficit_monthly_total.loc[month, building] = import_val-used_for_self

            # costi mensili pagati = deficit * price + system access mensile
            deficit_costs = deficit_monthly_total.loc[month, building] * price + monthly_system_access.loc[month, building]
            deficit_costs_monthly.loc[month, building] = deficit_costs

            # ✅ stesso valore (nome più “parlante”)
            #energy_costs_from_deficit.loc[month, building] = deficit * price + monthly_system_access.loc[month, building]

            self_consumed_energy.loc[month, building] = used_for_self
            energy_shared.loc[month, building] = shared_energy_val

            # self_consumed_total = PV production - unused credits (end of month)
            pv_prod_val = production_monthly.loc[month, building]
            PSC_kWh_val = PSC_monthly_kWh.loc[month, building]
            monthly_access_val = monthly_system_access.loc[month, building]

            pv_prod.loc[month, building] = pv_prod_val
            PSC_kWh.loc[month, building] = PSC_kWh_val
            monthly_access_charges.loc[month, building] = monthly_access_val
            self_consumed_total.loc[month, building] = max(pv_prod_val-shared_energy_val,0.0)
            self_consumed_onlyNM.loc[month, building] = max(pv_prod_val - shared_energy_val - PSC_kWh_val, 0.0)

            #self_consumed_total.loc[month, building] = max(pv_prod_val - unused_credits.loc[month, building], 0.0)

    return (credit_dist, balance_kWh, balance_costs,
            deficit_monthly_total, deficit_costs_monthly,
            monthly_average_price,
            self_consumed_energy, energy_shared, self_consumed_onlyNM,
            self_consumed_total,
            pv_prod, PSC_kWh, monthly_access_charges)


# Calculate monthly credits, deficit, unused credits
(credit_distribution, balance_kWh, balance_costs,
 deficit_monthly_total, deficit_costs_monthly,
 monthly_average_prices, self_consumed_energy,
 energy_shared, self_consumed_onlyNM, self_consumed_total, pv_prod,
 PSC_kWh, monthly_access_charges) = \
    calculate_credit_distribution_nopeak(import_monthly, export_monthly, monthly_prices, system_access_monthly,
                                         production_monthly, PSC_monthly_kWh)

# Create DataFrame for the final result of NM
with (pd.ExcelWriter(output_file_NM_nopeak, engine='openpyxl') as writer):
    consumption_monthly.to_excel(writer, sheet_name='Initial_demand')
    pv_prod.to_excel(writer, sheet_name='Initial_PV_production')
    PSC_kWh.to_excel(writer, sheet_name='Physical_selfcons')
    import_monthly.to_excel(writer, sheet_name='Import')
    export_monthly.to_excel(writer, sheet_name='Export')
    initial_costs_after_PSC.to_excel(writer, sheet_name='Initial_costs')
    #credit_distribution.to_excel(writer, sheet_name='Credit_distribution')
    balance_kWh.to_excel(writer, sheet_name='Balance_kWh')
    balance_costs.to_excel(writer, sheet_name='Balance_costs')
    deficit_costs_monthly.to_excel(writer, sheet_name='Deficit costs')
    deficit_monthly_total.to_excel(writer, sheet_name='Deficit_monthly_total')
    #energy_costs_from_deficit.to_excel(writer, sheet_name='Energy costs from deficit')
    monthly_average_prices.to_excel(writer, sheet_name='Monthly average price')
    #unused_credits.to_excel(writer, sheet_name='Unused_Credits')
    self_consumed_energy.to_excel(writer, sheet_name='Self_consumed energy')
    energy_shared.to_excel(writer, sheet_name='Energy shared')
    self_consumed_onlyNM.to_excel(writer, sheet_name='Self consumed onlyNM')
    self_consumed_total.to_excel(writer, sheet_name='Self_consumed_total')
    monthly_access_charges.to_excel(writer, sheet_name='Monthly access charges')
    initial_BAU_energy_costs.to_excel(writer, sheet_name='Initial BAU energy costs')