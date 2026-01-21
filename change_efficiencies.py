import shutil
import dbf
from simpledbf import Dbf5

base_path = r"C:\CityEnergyAnalyst\Paper_prova\Retrofit_sensitivity\inputs\building-properties"
typology_path = base_path + r"\typology.dbf"
supply_path   = base_path + r"\supply_systems.dbf"

# ---------- helper ----------
def norm(s):
    return str(s).strip()

def pick_field_case_insensitive(field_names, target):
    # ritorna il nome campo reale nel DBF, matchando senza distinzione maiuscole
    t = target.lower()
    for f in field_names:
        if f.lower() == t:
            return f
    return None

# ---------- backup ----------
backup_path = supply_path.replace(".dbf", "_BACKUP.dbf")
shutil.copy2(supply_path, backup_path)

# ---------- read typology ----------
typ = Dbf5(typology_path).to_dataframe()
# normalizza nomi colonne
typ.columns = [c.strip() for c in typ.columns]

# trova colonne Name e standard (anche se maiuscole)
typ_col_name = None
typ_col_std  = None
for c in typ.columns:
    if c.lower() == "name":
        typ_col_name = c
    if c.lower() == "standard":
        typ_col_std = c

if typ_col_name is None or typ_col_std is None:
    raise KeyError(f"Non trovo colonne 'Name' e/o 'standard' in typology.dbf. Colonne trovate: {list(typ.columns)}")

retro_standards = {"STANDARDAB_Retro2", "STANDARDMFH_Retro2"}

retro_names = set(
    norm(x).lower()
    for x in typ.loc[typ[typ_col_std].isin(retro_standards), typ_col_name]
)

print(f"Edifici in typology con standard Retro2: {len(retro_names)}")
print("Esempi retro (typology):", list(sorted(retro_names))[:5])

# ---------- open supply and update in-place ----------
table = dbf.Table(supply_path)
table.open(mode=dbf.READ_WRITE)

fields = list(table.field_names)
# trova i campi reali nel supply dbf
sup_name_f  = pick_field_case_insensitive(fields, "Name")
type_cs_f   = pick_field_case_insensitive(fields, "type_cs")
type_dhw_f  = pick_field_case_insensitive(fields, "type_dhw")
type_hs_f   = pick_field_case_insensitive(fields, "type_hs")

if sup_name_f is None:
    table.close()
    raise KeyError(f"Non trovo un campo 'Name' in supply_systems.dbf. Campi: {fields}")

missing = [x for x in [type_cs_f, type_dhw_f, type_hs_f] if x is None]
if missing:
    table.close()
    raise KeyError(f"Non trovo uno o più campi type_* in supply_systems.dbf. Campi: {fields}")

count = 0
seen_supply_names = []

for rec in table:
    n = norm(rec[sup_name_f]).lower()
    if len(seen_supply_names) < 5:
        seen_supply_names.append(n)

    if n in retro_names:
        with rec:
            rec[type_cs_f]  = "SUPPLY_COOLING_AS12"
            rec[type_dhw_f] = "SUPPLY_HOTWATER_AS12"
            rec[type_hs_f]  = "SUPPLY_HEATING_AS12"
        count += 1

table.close()

print("Esempi name (supply):", seen_supply_names)
print(f"Modificati {count} edifici")
print(f"Backup salvato in: {backup_path}")
