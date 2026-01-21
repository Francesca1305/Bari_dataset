from pathlib import Path
import pandas as pd
import numpy as np

# ====== INPUT / OUTPUT ======
base_dir = Path(r"C:\CityEnergyAnalyst\Paper_prova\Retrofit_sensitivity")
pv_dir = base_dir / "outputs/data/potentials/solar_PV5"
out_xlsx = base_dir / "PV_sensors_all_buildings.xlsx"
pv_totals_csv = base_dir / "outputs/data/potentials/solar_PV5/PV_PV5_total_buildings.csv"
# ====== COLONNE RICHIESTE (da PV_sensors) ======
cols = [
    "BUILDING","AREA_m2","tilt_deg","B_deg",
    "area_installed_module_m2"
]

rows = []

csv_files = sorted(pv_dir.glob("B*_PV_sensors.csv"))
if not csv_files:
    raise FileNotFoundError(f"Nessun file trovato in: {pv_dir}")

# ====== 1) 1 riga per edificio dai PV_sensors ======
for fp in csv_files:
    df = pd.read_csv(fp)

    # uniforma colonne mancanti
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA

    df = df[cols].copy()

    # numerici
    df["__aim"] = pd.to_numeric(df["area_installed_module_m2"], errors="coerce")
    df["__area"] = pd.to_numeric(df["AREA_m2"], errors="coerce")

    # pick riga rappresentativa
    if df["__aim"].notna().any() and (df["__aim"].fillna(0) > 0).any():
        pick = df.sort_values("__aim", ascending=False).iloc[0]
    else:
        pick = df.sort_values("__area", ascending=False).iloc[0]

    pick = pick.drop(labels=["__aim", "__area"], errors="ignore")
    rows.append(pick)

out = pd.DataFrame(rows)

# ordina per numero edificio (opzionale)
out["__Bnum"] = out["BUILDING"].astype(str).str.extract(r"B(\d+)").astype(float)
out = out.sort_values("__Bnum").drop(columns="__Bnum")

# assicura tipo numerico per area_installed_module_m2
out["area_installed_module_m2"] = pd.to_numeric(out["area_installed_module_m2"], errors="coerce")

# ====== 2) Leggi PV totals e prendi solo PV_roofs_top_E_kWh ======
pv_tot = pd.read_csv(pv_totals_csv)
pv_tot = pv_tot[["Name", "PV_roofs_top_E_kWh"]].copy()
pv_tot["Name"] = pv_tot["Name"].astype(str)
pv_tot["PV_roofs_top_E_kWh"] = pd.to_numeric(pv_tot["PV_roofs_top_E_kWh"], errors="coerce")

# merge
out = out.merge(pv_tot, left_on="BUILDING", right_on="Name", how="left").drop(columns=["Name"])

# ====== 3) Percentuali sul totale ======
total_pv = out["PV_roofs_top_E_kWh"].sum(skipna=True)
total_area_inst = out["area_installed_module_m2"].sum(skipna=True)

# evita divisione per zero
out["PV_roofs_top_E_kWh_pct_of_total"] = np.where(
    total_pv > 0,
    out["PV_roofs_top_E_kWh"] / total_pv * 100.0,
    np.nan
)

out["area_installed_module_m2_pct_of_total"] = np.where(
    total_area_inst > 0,
    out["area_installed_module_m2"] / total_area_inst * 100.0,
    np.nan
)

# (opzionale) arrotonda a 2 decimali
out["PV_roofs_top_E_kWh_pct_of_total"] = out["PV_roofs_top_E_kWh_pct_of_total"].round(2)
out["area_installed_module_m2_pct_of_total"] = out["area_installed_module_m2_pct_of_total"].round(2)

# ====== OUTPUT ======
out.to_excel(out_xlsx, index=False)
print(f"Creato: {out_xlsx}")
print(f"Totale PV_roofs_top_E_kWh = {total_pv:.2f}")
print(f"Totale area_installed_module_m2 = {total_area_inst:.2f}")

