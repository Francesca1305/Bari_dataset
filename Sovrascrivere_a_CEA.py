import pandas as pd
import os

# Cartelle
output_folder = "Data/Building_schedules_CEA"
cea_folder = ("C:/CityEnergyAnalyst/Paper_prova/BAU_scenario_prova3/"
              "inputs/building-properties/schedules")

# Assicura che la cartella CEA esista
if not os.path.exists(cea_folder):
    print("❌ La cartella 'Data/CEA' non esiste. Interruzione.")
    exit()

# Processa ogni file in output_folder
for file in os.listdir(output_folder):
    if file.endswith(".csv"):
        building_id = file.split(".csv")[0]  # Estrarre l'ID dell'edificio
        output_path = os.path.join(output_folder, file)
        cea_path = os.path.join(cea_folder, file)  # Stesso nome in CEA

        # Verifica che il file esista in CEA
        if os.path.exists(cea_path):
            # Leggi i dati aggiornati
            updated_df = pd.read_csv(output_path, skiprows=2)  # Ignora METADATA e MONTHLY_MULTIPLIER

            # Leggi il file originale di CEA
            with open(cea_path, "r") as f:
                lines = f.readlines()

            cea_metadata = lines[:2]  # Mantiene le prime due righe
            cea_df = pd.read_csv(cea_path, skiprows=2)  # Salta METADATA e MULTIPLIER

            # Sostituisci solo le colonne richieste
            for col in ["OCCUPANCY", "APPLIANCES", "LIGHTING", "WATER"]:
                if col in cea_df.columns and col in updated_df.columns:
                    cea_df[col] = updated_df[col]

            # Riscrivi il file originale con i nuovi dati
            with open(cea_path, "w", newline="") as f:
                f.writelines(cea_metadata)  # Scrive METADATA e MULTIPLIER
                cea_df.to_csv(f, index=False, float_format="%.1f")  # Scrive il nuovo contenuto

            print(f"🔄 Aggiornato: {cea_path}")
        else:
            print(f"⚠️ File {cea_path} non trovato. Saltato.")

print("✅ Sostituzione completata!")