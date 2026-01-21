import pandas as pd
import numpy as np
from pathlib import Path

name_scenario = "Retrofit_II_0.20PV"

# Percorso file Excel
# community_file = Path(
#     f"C:/CityEnergyAnalyst/Paper_prova/{name_scenario}/outputs/"
#     f"Elaboration_REC/PV5_20%_efficiency/community_bybuilding_{name_scenario}.xlsx"
# )

community_file=Path(r"C:\Users\franc\Desktop\ABM Bari\Elaboration_REC\community_bybuilding_BAU_scenario_prova3_stochastic.xlsx")

# =========================
# Lettura fogli necessari
# =========================
valutazione_CER_df = pd.read_excel(community_file, sheet_name="valutazione CER")
import_df = pd.read_excel(community_file, sheet_name="Import_kWh")
export_df = pd.read_excel(community_file, sheet_name="Export_kWh")

'''eventualmente, aggiungere revenues con meccanismo Rid'''


# Lista edifici (tutte le colonne tranne Date)
building_ids = [c for c in import_df.columns if c != "Date"]

# =========================
# Allineamento Date
# =========================
valutazione_CER_df['Date'] = pd.to_datetime(valutazione_CER_df['Date'])
import_df['Date'] = pd.to_datetime(import_df['Date'])
export_df['Date'] = pd.to_datetime(export_df['Date'])

# =========================
# Estrazione CSC ed Export
# =========================
CSC = valutazione_CER_df['CSC'].values.reshape(-1, 1)   # (8760, 1)
incentive = valutazione_CER_df['Incentive'].values.reshape(-1, 1)
price_purchase = valutazione_CER_df['Price purchase'].values.reshape(-1, 1)
RiD = valutazione_CER_df['Price surplus'].values.reshape(-1, 1)
Revenues_surplus_after_REC = valutazione_CER_df['Energy revenues REC_surplus'].values.reshape(-1, 1)
import_values = import_df[building_ids].values      # (8760, n_buildings)
export_values = export_df[building_ids].values

# =========================
# Calcolo quote di export
# =========================
import_sum = import_values.sum(axis=1).reshape(-1, 1)
export_sum = export_values.sum(axis=1).reshape(-1, 1)

# Evita divisione per zero
export_shares = np.zeros_like(export_values)
np.divide(
    export_values,
    export_sum,
    out=export_shares,
    where=export_sum > 0
)
import_shares = np.zeros_like(import_values)
np.divide(
    import_values,
    import_sum,
    out=import_shares,
    where=import_sum > 0
)

# =========================
# Distribuzione CSC
# =========================
Energy_costs_import = import_values*price_purchase
CSC_distributed_import = CSC * import_shares
CSC_distributed_export = CSC * export_shares
CSC_distributed_revenues = CSC * export_shares * incentive
CSC_distributed_revenues_withRid = CSC * export_shares * (incentive + RiD)
revenues_after_REC_distributed = Revenues_surplus_after_REC * export_shares

# =========================
# Creazione DataFrame finale
# =========================
distribution_CSC_import_df = pd.DataFrame(CSC_distributed_import, columns=building_ids)
distribution_CSC_import_df.insert(0, 'Date', valutazione_CER_df['Date'])
distribution_CSC_export_df = pd.DataFrame(CSC_distributed_export, columns=building_ids)
distribution_CSC_export_df.insert(0, 'Date', valutazione_CER_df['Date'])
distribution_CSC_revenues_df = pd.DataFrame(CSC_distributed_revenues, columns=building_ids)
distribution_CSC_revenues_df.insert(0, 'Date', valutazione_CER_df['Date'])
distribution_CSC_revenues_withRiD_df = pd.DataFrame(CSC_distributed_revenues_withRid, columns=building_ids)
distribution_CSC_revenues_withRiD_df.insert(0, 'Date', valutazione_CER_df['Date'])
energy_costs_import_df = pd.DataFrame(Energy_costs_import, columns=building_ids)
energy_costs_import_df.insert(0, 'Date', valutazione_CER_df['Date'])
distribution_revenues_after_REC_df = pd.DataFrame(revenues_after_REC_distributed, columns=building_ids)
distribution_revenues_after_REC_df.insert(0, 'Date', valutazione_CER_df['Date'])

# =========================
# Scrittura nuovo sheet
# =========================
with pd.ExcelWriter(community_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    distribution_CSC_import_df.to_excel(writer, sheet_name="CSC_import", index=False)
    distribution_CSC_export_df.to_excel(writer, sheet_name="CSC_export", index=False)
    distribution_CSC_revenues_df.to_excel(writer, sheet_name="CSC_rev_TIP", index=False)
    distribution_CSC_revenues_withRiD_df.to_excel(writer, sheet_name="CSC_rev_TIP_RiD", index=False)
    energy_costs_import_df.to_excel(writer, sheet_name="Import_costs", index=False)
    distribution_revenues_after_REC_df.to_excel(writer, sheet_name="REC_surplus_rev", index=False)


print("Sheet creato correttamente.")
