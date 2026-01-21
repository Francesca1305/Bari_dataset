import pandas as pd
import os

# Input e output directory
input_folder = "Data/Building_profiles_all"
output_folder = "Data/Building_schedules_CEA"
os.makedirs(output_folder, exist_ok=True)

# Struttura header CEA
metadata = ["METADATA", "CH-SIA-2014", "SINGLE_RES"] + [""] * 8
#monthly_multiplier = ["MONTHLY_MULTIPLIER"] + [0.8] * 12
monthly_multiplier = ["MONTHLY_MULTIPLIER"] + [1, 1, 1, 1, 1, 1, 1, 0.8, 1, 1, 1, 1]
# Giorni della settimana e ore
days = ["WEEKDAY"] * 24 + ["SATURDAY"] * 24 + ["SUNDAY"] * 24
hours = list(range(1, 25)) * 3

# Processa ogni file esistente
for file in os.listdir(input_folder):
    if file.endswith(".xlsx"):
        building_id = file.split(".xlsx")[0]
        input_path = os.path.join(input_folder, file)
        output_path = os.path.join(output_folder, f"B{building_id}.csv")

        # Leggi i valori "Average" per ogni foglio
        average_values = {}
        with pd.ExcelFile(input_path) as xls:
            for category in ["occupancy", "appliances", "lighting", "DHW"]:
                if category in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=category)
                    avg_row = df[df.iloc[:, 0] == "Average"].iloc[:, 1:].values.flatten()

                    # Assicura che ci siano 72 valori correttamente distribuiti
                    if len(avg_row) == 72:
                        average_values[category] = avg_row
                    else:
                        print(f"⚠️ Warning: {category} in {building_id}.xlsx ha "
                              f"{len(avg_row)} valori invece di 72.")
                        average_values[category] = [0] * 72

        # Crea DataFrame finale con la distribuzione corretta dei valori
        schedule_data = pd.DataFrame({
            "OCCUPANCY": list(average_values["occupancy"]),
            "APPLIANCES": list(average_values["appliances"]),
            "LIGHTING": list(average_values["lighting"]),
            "SERVERS": [0] * 72,
            "WATER": list(average_values["DHW"]),
            "HEATING": ["SETPOINT", "SETPOINT", "SETPOINT", "SETPOINT", "SETPOINT", "SETPOINT",
                        "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK",
                        "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK",
                        "SETPOINT", "SETPOINT", "SETPOINT"] * 3,
            "COOLING": ["SETPOINT", "SETPOINT", "SETPOINT", "SETPOINT", "SETPOINT", "SETPOINT",
                        "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK",
                        "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK", "SETBACK",
                        "SETPOINT", "SETPOINT", "SETPOINT"] * 3,
            "PROCESSES": [0] * 72,
            "ELECTROMOBILITY": [0] * 72,
            "DAY": days,
            "HOUR": hours
        })

        # Salva in CSV
        with open(output_path, "w", newline="") as f:
            # Scrive la riga METADATA
            f.write(",".join(map(str, metadata)) + "\n")
            # Scrive la riga MONTHLY_MULTIPLIER
            f.write(",".join(map(str, monthly_multiplier)) + "\n")
            # Scrive i dati
            schedule_data.to_csv(f, index=False, float_format="%.1f")

        print(f"✅ Creato file: {output_path}")

print("🎉 Elaborazione completata!")



