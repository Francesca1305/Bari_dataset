import pandas as pd
from pathlib import Path

base_path = Path("C:\CityEnergyAnalyst\Paper_prova")
INPUT_XLSX = base_path/"Validation_consumption\outputs\Elaboration_REC\community_bybuilding_Validation_consumption.xlsx"
#INPUT_XLSX = base_path/"Elaboration_NG\community_bybuilding_BAU_scenario_prova3_NG.xlsx"
SHEET_NAME = "Demand_kWh"
#SHEET_NAME = "Heating_NG"
OUTPUT_XLSX = base_path/"monthly_demand_aggregation_EE.xlsx"

# Leggi il foglio
df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_NAME)

# Assicurati che Date sia datetime
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

# Tieni solo righe con Date valide
df = df.dropna(subset=["Date"]).copy()

# Converti tutte le colonne edifici in numerico (se ci sono stringhe/spazi)
building_cols = [c for c in df.columns if c != "Date"]
df[building_cols] = df[building_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

# Aggregazione mensile: somma dei kWh per mese
monthly = (
    df.set_index("Date")[building_cols]
      .resample("MS")          # Month Start: 2016-01-01, 2016-02-01, ...
      .sum()
)

# Aggiungi una colonna "Month" più leggibile (opzionale)
monthly_out = monthly.copy()
monthly_out.insert(0, "Month", monthly_out.index.to_period("M").astype(str))
monthly_out = monthly_out.reset_index(drop=True)

# Salva su Excel
with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
    monthly_out.to_excel(writer, sheet_name="Monthly_kWh", index=False)

print(f"Creato: {OUTPUT_XLSX} (foglio: Monthly_kWh)")
