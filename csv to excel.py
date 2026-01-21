import pandas

import pandas as pd

# Leggi il file CSV
df_deterministic = pd.read_csv('C:/CityEnergyAnalyst/Paper_prova/BAU_scenario_prova3/outputs/data/'
                               'demand_semi deterministic/B17.csv')
df_stochastic = pd.read_csv('C:/CityEnergyAnalyst/Paper_prova/BAU_scenario_prova3/outputs/data/'
                               'demand_stochastic/B17.csv')

# Salva come file Excel (.xlsx)
df_deterministic.to_excel('B17_D_det.xlsx', index=False)
df_stochastic.to_excel('B17_D_stoc.xlsx', index=False)