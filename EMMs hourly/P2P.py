import pandas as pd
import numpy as np
from pathlib import Path

def calculate_dynamic_price(demand, generation, import_grid,
                            export_generation, base_price, alpha, k, wholesale_price, access_charges):
    # Substitute values = 0 in the generation
    #generation = np.where(generation == 0, np.nan, generation)
    # Calculate the hourly Supply Demand Ratio
    SDR = generation/demand
    # Calculate Demand Peak (D_peak) over the whole year
    D_peak = demand.max()
    # Calculate the Peak Demand factor
    PD_factor = 1 + (alpha*(D_peak - demand) / D_peak)
    # Calculate the IDP
    IDP = np.where(generation > 0, base_price * (1 + k * (1 - SDR)
                                                 * PD_factor), base_price)
    # IDP cannot be negative and the minimum value is 0.01
    IDP = np.maximum(IDP, wholesale_price)
    access_charges = access_charges
    export_generation = np.nan_to_num(export_generation, nan=0.0)
    import_grid = np.nan_to_num(import_grid, nan=0.0)
    # Calculate hourly Collective Self Consumption
    collective_SC = np.where((import_grid > 0) & (export_generation > 0),
                   np.minimum(import_grid, export_generation),
                   0)
    # Costs of CSC with IDP
    costs_CSC_access = np.where(collective_SC > 0, access_charges, 0.0)
    costs_collective_SC = collective_SC * IDP + costs_CSC_access

    # Calculate the eventual surplus and deficit
    surplus_REC = np.maximum(export_generation - collective_SC, 0)
    deficit_REC = np.maximum(import_grid - collective_SC, 0)
    surplus_REC_revenues = surplus_REC * wholesale_price
    deficit_REC_costs = deficit_REC * base_price + access_charges

    return (SDR, PD_factor, IDP, access_charges, collective_SC, costs_collective_SC,
            surplus_REC, deficit_REC, surplus_REC_revenues, deficit_REC_costs)


# =========================
# INPUT
# =========================
community_file = Path(
    r"C:\Users\franc\Desktop\ABM Bari\Elaboration_REC\community_bybuilding_BAU_scenario_prova3_stochastic.xlsx"
)

# =========================
# Read valutazione CER
# =========================
valutazione_CER_df = pd.read_excel(community_file, sheet_name="valutazione CER")
import_df = pd.read_excel(community_file, sheet_name="Import_kWh")
export_df = pd.read_excel(community_file, sheet_name="Export_kWh")

valutazione_CER_df['Date'] = pd.to_datetime(valutazione_CER_df['Date'])
#valutazione_CER_df.reset_index('Date', inplace=True)
import_df['Date'] = pd.to_datetime(import_df['Date'])
export_df['Date'] = pd.to_datetime(export_df['Date'])
building_ids = [c for c in import_df.columns if c != "Date"]
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
# Extract variables
# =========================
demand = valutazione_CER_df['total_cons'].values
generation = valutazione_CER_df['total_PV'].values
import_grid = valutazione_CER_df['import'].values
export_generation = valutazione_CER_df['export'].values
base_price = valutazione_CER_df['Price purchase'].values
wholesale_price = valutazione_CER_df['Price surplus'].values

# System access charges
# 👉 se non li hai orari, puoi usare un valore fisso
access_charges = 0  # oppure €/kWh costante

# =========================
# IDP parameters
# =========================
alpha = 0.3
k = 0.5

# =========================
# Apply IDP
# =========================
(SDR, PD_factor, IDP, access_charges, collective_SC,
 costs_collective_SC, surplus_REC, deficit_REC,
 surplus_REC_revenues, deficit_REC_costs) = (
    calculate_dynamic_price(
        demand,
        generation,
        import_grid,
        export_generation,
        base_price,
        alpha,
        k,
        wholesale_price,
        access_charges
    )
)

# =========================
# Results DataFrame
# =========================
results_IDP = pd.DataFrame({
    "Date": valutazione_CER_df.index,
    "Demand": demand,
    "Generation": generation,
    "Import": import_grid,
    "Export": export_generation,
    "Base_price": base_price,
    "Wholesale_price": wholesale_price,
    "SDR": SDR,
    "PD_factor": PD_factor,
    "IDP": IDP,
    "Collective_SC": collective_SC,
    "Costs_Collective_SC_IDP": costs_collective_SC,
    "Surplus_REC": surplus_REC,
    "Deficit_REC": deficit_REC,
    "Surplus_REC_revenues": surplus_REC_revenues,
    "Deficit_REC_costs": deficit_REC_costs
})

results_IDP = results_IDP.round(4)

# =========================
# Distribuzione CSC
# =========================
collective_SC = collective_SC.reshape(-1, 1)
surplus_REC_revenues = surplus_REC_revenues.reshape(-1, 1)
base_price = base_price.reshape(-1, 1)
IDP = IDP.reshape(-1, 1)
CSC_distributed_import = collective_SC * import_shares
CSC_distributed_export = collective_SC * export_shares
revenues_after_REC_distributed = surplus_REC_revenues * export_shares
Energy_costs_IDP_import = CSC_distributed_import*IDP + (import_values-CSC_distributed_import)*base_price
IDP_distributed_revenues = (collective_SC * export_shares) * IDP

# =========================
# Save to Excel
# =========================
output_IDP_file = Path(
    r"C:\Users\franc\Desktop\ABM Bari\Elaboration_REC\P2P_with_IDP.xlsx"
)

# =========================
# Creazione DataFrame finale
# =========================
distribution_CSC_import_df = pd.DataFrame(CSC_distributed_import, columns=building_ids)
distribution_CSC_import_df.insert(0, 'Date', valutazione_CER_df['Date'])
distribution_CSC_export_df = pd.DataFrame(CSC_distributed_export, columns=building_ids)
distribution_CSC_export_df.insert(0, 'Date', valutazione_CER_df['Date'])
distribution_IDP_revenues_df = pd.DataFrame(IDP_distributed_revenues, columns=building_ids)
distribution_IDP_revenues_df.insert(0, 'Date', valutazione_CER_df['Date'])
energy_costs_IDP_import_df = pd.DataFrame(Energy_costs_IDP_import, columns=building_ids)
energy_costs_IDP_import_df.insert(0, 'Date', valutazione_CER_df['Date'])
distribution_revenues_after_REC_df = pd.DataFrame(revenues_after_REC_distributed, columns=building_ids)
distribution_revenues_after_REC_df.insert(0, 'Date', valutazione_CER_df['Date'])

# =========================
# Scrittura nuovo sheet
# =========================

with pd.ExcelWriter(output_IDP_file, engine="openpyxl", mode="w") as writer:
    results_IDP.to_excel(writer, sheet_name="IDP_community", index=False)
    distribution_CSC_import_df.to_excel(writer, sheet_name="CSC_import", index=False)
    distribution_CSC_export_df.to_excel(writer, sheet_name="CSC_export", index=False)
    distribution_IDP_revenues_df.to_excel(writer, sheet_name="CSC_rev_IDP", index=False)
    energy_costs_IDP_import_df.to_excel(writer, sheet_name="Import_costs_IDP", index=False)
    distribution_revenues_after_REC_df.to_excel(writer, sheet_name="REC_surplus_rev", index=False)

# IDP = "P2P_with_IDP.xlsx"
# with pd.ExcelWriter(IDP, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
#     results_IDP.to_excel(writer, sheet_name="IDP_community", index=False)

print("IDP community sheet creato correttamente.")
