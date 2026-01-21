import pandas as pd
import json
import random
import os
from collections import Counter

# File paths
census_data_file = "Data/census_data_out.json"
building_data_file = "Data/Shapefile_27.12_con census.json"
profile_files = {
    "occupancy": "Data/Building_profiles/profiles_HH_ncomp_occupancy.xlsx",
    "appliances": "Data/Building_profiles/profiles_HH_ncomp_Appliances.xlsx",
    "lighting": "Data/Building_profiles/profiles_HH_ncomp_Lighting.xlsx",
    "DHW": "Data/Building_profiles/profiles_HH_ncomp_DHW.xlsx"
}
output_folder = "Data/Building_profiles_all_27.12"

# Household types mapping (sheet names → simplified codes)
ncomp_types = {
    "1 ncomp, occupied": "1_comp_work",
    "1 ncomp, retired": "1_comp_ret",
    "2 ncomp, occupied": "2_comp_work",
    "2 ncomp, retired": "2_comp_ret",
    "3 members": "3_comp",
    "More": "4_comp_more"
}

# Ensure output folder exists
os.makedirs(output_folder, exist_ok=True)

# Load census and building data
with open(census_data_file, "r") as f:
    census_data = json.load(f)
with open(building_data_file, "r") as f:
    building_data = json.load(f)

# Load all profiles into a dictionary
profiles = {}
for category, file_path in profile_files.items():
    profiles[category] = {
        ncomp_types[sheet]: pd.read_excel(file_path, sheet_name=sheet).values.tolist()
        for sheet in ncomp_types
    }

building_household_summary = []

def assign_households_to_residential_buildings(census_sections, building_data):
    building_summary = {}
    building_estimated_residents = {}

    for feature in census_sections["features"]:
        census = feature["properties"]
        census_id = census["SEZ21"]
        census_population = census["total resident population"]
        total_census_households = census["total households"]
        no_study_titles = census["P86"]
        elementary_title = census["P87"]
        middle_school_title = census["P88"]
        secondary_school_title = census["P89"]
        university_title = census["P90"]
        unknown_title = census["unknown_education"]


        if census_population <= 0 or total_census_households <= 0:
            continue

        census_occupied = (
            (census["Italian occupied_IT10"] + census["Foreign occupied_ST31"])
            / census["total resident population"]
        )

        residential_buildings = [
            b["properties"] for b in building_data["features"]
            if b["properties"]["SEZ21"] == census_id and b["properties"]["function"] in [11, 12]
        ]
        if not residential_buildings:
            continue

        # --- Calculate heated area and assign residents ---
        total_heated_area = sum(b["Area"] * 0.82 * b["nfloors"] for b in residential_buildings)
        avg_area_per_person = total_heated_area / census_population if census_population > 0 else 1

        for b in residential_buildings:
            b["heated_area"] = b["Area"] * 0.82 * b["nfloors"]
            b["estimated_residents"] = max(1, round(b["heated_area"] / avg_area_per_person))

        # --- Adjust population to match census ---
        assigned_population = sum(b["estimated_residents"] for b in residential_buildings)
        diff = census_population - assigned_population
        for _ in range(abs(diff)):
            choice = random.choice(residential_buildings)
            choice["estimated_residents"] += 1 if diff > 0 else -1

        # --- Create household pool for this census section ---
        households = []
        hh_distribution = [
            ("1_comp_work", int(census["HH_1 comp"] * census_occupied)),
            ("1_comp_ret", int(census["HH_1 comp"] * (1 - census_occupied))),
            ("2_comp_work", int(census["HH_2 comp"] * census_occupied)),
            ("2_comp_ret", int(census["HH_2 comp"] * (1 - census_occupied))),
            ("3_comp", int(census["HH_3 comp"])),
            ("4_comp_more", int(census["HH_4 comp"] + census["HH_5 comp"] + census["HH_6 comp or more"]))
        ]

        for ncomp_type, count in hh_distribution:
            for _ in range(count):
                household_profiles = {cat: profiles[cat][ncomp_type] for cat in profile_files}
                households.append((ncomp_type, household_profiles))

        random.shuffle(households)
        print(f"census section ID {census_id}: {avg_area_per_person} m2/person")

        # --- Assign to buildings ---
        for building in residential_buildings:
            building_id = building["ID"]
            num_households = max(
                1,
                round(building["estimated_residents"] / census_population * total_census_households)
            )

            summary = {
                "single_worker": {"Wasteful": 0, "Average": 0, "Saver": 0},
                "single_retired": {"Wasteful": 0, "Average": 0, "Saver": 0},
                "couple_workers": {"Wasteful": 0, "Average": 0, "Saver": 0},
                "couple_retired": {"Wasteful": 0, "Average": 0, "Saver": 0},
                "families": {"Wasteful": 0, "Average": 0, "Saver": 0},
            }

            assigned_profiles = {cat: [] for cat in profile_files}
            building_hh_count = {}

            for _ in range(num_households):
                if not households:
                    break
                ncomp_type, household_profiles = households.pop(0)

                # count households
                building_hh_count[ncomp_type] = building_hh_count.get(ncomp_type, 0) + 1

                # update summary
                if ncomp_type == "1_comp_work":
                    summary["single_worker"]["Average"] += 1
                elif ncomp_type == "1_comp_ret":
                    summary["single_retired"]["Average"] += 1
                elif ncomp_type == "2_comp_work":
                    summary["couple_workers"]["Average"] += 1
                elif ncomp_type == "2_comp_ret":
                    summary["couple_retired"]["Average"] += 1
                elif ncomp_type in ["3_comp", "4_comp_more"]:
                    summary["families"]["Average"] += 1

                # assign hourly profiles
                for cat in profile_files:
                    assigned_profiles[cat].append(
                        [ncomp_type] + [val[0] if isinstance(val, list) else val for val in household_profiles[cat]]
                    )

            building_summary[building_id] = summary
            # Calculate number of occupied residents from household summary
            occupied_count = (
                summary["single_worker"]["Average"] * 1 +
                summary["couple_workers"]["Average"] * 2 +
                summary["families"]["Average"] * 2
            )

            # --- Education data for the census section ---
            section_education = {
                "no_study": no_study_titles,
                "elementary": elementary_title,
                "middle_school": middle_school_title,
                "secondary_school": secondary_school_title,
                "university": university_title,
                "unknown_edu": unknown_title
            }
            # --- RANDOM education allocation per building ---

            # Calcola le probabilità di ciascun livello di istruzione nella sezione
            total_section_education = sum(section_education.values())
            if total_section_education == 0:
                total_section_education = 1

            education_levels = list(section_education.keys())
            education_probs = [
                value / total_section_education for value in section_education.values()
            ]

            # Estrai casualmente un titolo per ogni residente stimato nell’edificio
            random_education = random.choices(
                population=education_levels,
                weights=education_probs,
                k=building["estimated_residents"]
            )

            # Conta quanti residenti per ciascun livello
            education_count = dict(Counter(random_education))

            # Aggiungi eventuali livelli mancanti (per mantenere la chiave anche se 0)
            for level in section_education.keys():
                education_count.setdefault(level, 0)

            # --- LOCAL REBALANCING per garantire coerenza con estimated_residents ---

            sum_edu = sum(education_count.values())
            if sum_edu < building["estimated_residents"]:
                diff = building["estimated_residents"] - sum_edu
                # aggiungi diff residenti casuali a livelli esistenti
                for _ in range(diff):
                    add_level = random.choice(list(education_count.keys()))
                    education_count[add_level] += 1
            elif sum_edu > building["estimated_residents"]:
                diff = sum_edu - building["estimated_residents"]
                # rimuovi diff residenti casuali dove il conteggio > 0
                for _ in range(diff):
                    candidates = [lvl for lvl, val in education_count.items() if val > 0]
                    if candidates:
                        rem_level = random.choice(candidates)
                        education_count[rem_level] -= 1

            # --- Estimate households and residents per type ---
            hh_types_residents = {
                "single_worker": summary["single_worker"]["Average"] * 1,
                "single_retired": summary["single_retired"]["Average"] * 1,
                "couple_workers": summary["couple_workers"]["Average"] * 2,
                "couple_retired": summary["couple_retired"]["Average"] * 2,
                "families": summary["families"]["Average"] * 3
            }
            total_building_residents = sum(hh_types_residents.values())
            if total_building_residents == 0:
                total_building_residents = 1

            # --- Build a list of "virtual residents" for this building ---
            education_levels = []
            for level, count in education_count.items():
                education_levels.extend([level] * count)

            # Se il numero di residenti stimati > persone con titolo assegnato, riempi con "unknown"
            #if len(education_levels) < building["estimated_residents"]:
            #    education_levels.extend(["unknown"] * (building["estimated_residents"] - len(education_levels)))

            # Shuffle for randomness
            random.shuffle(education_levels)

            # --- Assign education levels to households randomly ---
            start_idx = 0
            household_education_detail = {}
            for hh_type, nres in hh_types_residents.items():
                assigned = education_levels[start_idx:start_idx + nres] if nres > 0 else []
                start_idx += nres
                edu_counter = Counter(assigned)
                household_education_detail[hh_type] = {
                    "list": assigned,
                    "count": dict(edu_counter)
                }

            # --- Aggregate education counts per building ---
            aggregated_education = {}
            for level in section_education.keys():
                total = 0
                for hh in household_education_detail.values():
                    if isinstance(hh, dict) and "count" in hh:
                        total += hh["count"].get(level, 0)
                aggregated_education[level] = total

            # --- FINAL LOCAL REBALANCING ---
            # Garantisce coerenza tra estimated_residents, education e household_types

            sum_edu = sum(aggregated_education.values())
            sum_household = sum(
                sum(v.get("count", {}).values()) for v in household_education_detail.values()
            )

            # Se la somma dell’educazione o dei tipi household è inferiore a estimated_residents
            if building["estimated_residents"] > max(sum_edu, sum_household):
                diff = building["estimated_residents"] - max(sum_edu, sum_household)
                for _ in range(diff):
                    add_level = random.choice(list(aggregated_education.keys()))
                    aggregated_education[add_level] += 1

            # Se invece eccedono
            elif building["estimated_residents"] < min(sum_edu, sum_household):
                diff = min(sum_edu, sum_household) - building["estimated_residents"]
                for _ in range(diff):
                    candidates = [lvl for lvl, val in aggregated_education.items() if val > 0]
                    if candidates:
                        rem_level = random.choice(candidates)
                        aggregated_education[rem_level] -= 1

            # Ora allinea anche la somma di household_types (se manca qualche individuo)
            sum_household_final = sum(
                sum(v.get("count", {}).values()) for v in household_education_detail.values()
            )

            if sum_household_final < building["estimated_residents"]:
                diff = building["estimated_residents"] - sum_household_final
                for _ in range(diff):
                    hh_choice = random.choice(list(household_education_detail.keys()))
                    lvl_choice = random.choice(list(aggregated_education.keys()))
                    household_education_detail[hh_choice]["count"][lvl_choice] = (
                        household_education_detail[hh_choice]["count"].get(lvl_choice, 0) + 1
                    )

            # --- Save to dictionary ---
            building_estimated_residents[building_id] = {
                "estimated_residents": building["estimated_residents"],
                "occupied": occupied_count,
                "education": aggregated_education,
                "household_types": {
                    hh_type: hh["count"] for hh_type, hh in household_education_detail.items()
                }
            }

            # save detailed Excel
            if any(assigned_profiles.values()):
                file_path = os.path.join(output_folder, f"{building_id}.xlsx")
                with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
                    for cat, data in assigned_profiles.items():
                        if not data:
                            continue
                        col_names = ["Household Type"] + [f"Hour_{i+1}" for i in range(len(data[0]) - 1)]
                        df = pd.DataFrame(data, columns=col_names)
                        avg_row = ["Average"] + df.iloc[:, 1:].mean().tolist()
                        df.loc[len(df)] = avg_row
                        df.to_excel(writer, sheet_name=cat, index=False)
                print(f"File creato: {file_path}")

            # --- GLOBAL REBALANCING per la SEZIONE CENSUARIA ---

            # # Calcola la somma di education per tutti gli edifici di questa sezione
            # section_buildings = [bid for bid in building_estimated_residents if
            #                      building_estimated_residents[bid].get("Census_Section", census_id) == census_id]
            # section_sum = {lvl: 0 for lvl in section_education.keys()}
            # for bid in section_buildings:
            #     edu_dict = building_estimated_residents[bid]["education"]
            #     for lvl, val in edu_dict.items():
            #         section_sum[lvl] += val
            #
            # # Differenze con i valori censuari
            # edu_levels = list(section_education.keys())
            # for lvl in edu_levels:
            #     diff = section_education[lvl] - section_sum[lvl]
            #     if diff == 0:
            #         continue
            #
            #     # Se mancano residenti per un livello → aggiungili randomicamente
            #     if diff > 0:
            #         for _ in range(diff):
            #             target_building = random.choice(section_buildings)
            #             building_estimated_residents[target_building]["education"][lvl] += 1
            #     # Se sono in eccesso → rimuovili dove presenti
            #     else:
            #         for _ in range(abs(diff)):
            #             candidates = [
            #                 bid for bid in section_buildings
            #                 if building_estimated_residents[bid]["education"].get(lvl, 0) > 0
            #             ]
            #             if candidates:
            #                 target_building = random.choice(candidates)
            #                 building_estimated_residents[target_building]["education"][lvl] -= 1

            # Save summary for later aggregation
            for hh_type, count in building_hh_count.items():
                building_household_summary.append({
                    "Building_ID": building_id,
                    "Census_Section": census_id,
                    "Household_Type": hh_type,
                    "Count": count
                })

    # ---- SAVE ONE JSON FILE AT THE END ----
    output_json_path = os.path.join(output_folder, "building_household_summary.json")
    with open(output_json_path, "w") as f:
        json.dump(building_summary, f, indent=2)
    print(f"✅ Saved household summary JSON to {output_json_path}")

    residents_json_path = os.path.join(output_folder, "building_estimated_residents.json")
    with open(residents_json_path, "w") as f:
        json.dump(building_estimated_residents, f, indent=2)

# ---- RUN PROCESS ----
assign_households_to_residential_buildings(census_data, building_data)

summary_df = pd.DataFrame(building_household_summary)
pivot_df = summary_df.pivot_table(
    index=["Census_Section", "Building_ID"],
    columns="Household_Type",
    values="Count",
    aggfunc="sum",
    fill_value=0
).reset_index()

pivot_df = pivot_df.sort_values(by=["Census_Section", "Building_ID"])
summary_file_path = os.path.join(output_folder, "Household_assignment_summary.xlsx")
pivot_df.to_excel(summary_file_path, index=False)
print(f"\n✅ File riepilogativo creato: {summary_file_path}")
