import pandas as pd
import json
import random
import os

# File paths
census_data = "Data/census_data_out.json"
building_data = "Data/Building_with_census_updated.json"
profile_file = "Data/profiles_HH.xlsx"
output_folder = "Data/Building_profiles"

# Household types mapping
ncomp_types = {
    "1component_worker": "1_comp_work",
    "1component_retired": "1_comp_ret",
    "2components_working": "2_comp_work",
    "2components_retired": "2_comp_ret",
    "3components": "3_comp",
    "4components_more": "4_comp_more"
}

# Ensure output folder exists
os.makedirs(output_folder, exist_ok=True)

# Load data
with open(census_data, 'r') as f:
    census_data = json.load(f)

with open(building_data, 'r') as f:
    building_data = json.load(f)

profiles = {
    ncomp_types[sheet]: pd.read_excel(profile_file, sheet_name=sheet).values.tolist()
    for sheet in ncomp_types
}


def assign_households_to_residential_buildings(census_sections, building_data):
    building_summary = {}  # <-- define here
    building_estimated_residents = {}

    for feature in census_sections["features"]:
        census = feature["properties"]
        census_id = census["SEZ21"]
        census_population = census["total resident population"]
        total_census_households = census["total households"]
        census_occupied = ((census["Italian occupied_IT10"] +
                            census["Foreign occupied_ST31"]) /
                           census["total resident population"]) if census_population > 0 else 0

        residential_buildings = [b["properties"] for b in building_data["features"]
                                 if b["properties"]["SEZ21"] == census_id and b["properties"]["function"] in [11, 12]]
        if not residential_buildings:
            continue

        # Calculate total heated area for residential buildings
        total_heated_area = sum(b["Shape_Area"] * b["nfloors"] for b in residential_buildings)
        avg_area_per_person = total_heated_area / census_population if census_population > 0 else 1

        # Assign estimated residents to each building
        for building in residential_buildings:
            building["heated_area"] = building["Shape_Area"] * building["nfloors"]
            building["estimated_residents"] = max(1, round(building["heated_area"] / avg_area_per_person))

        # Adjust total residents to match census population
        assigned_population = sum(b["estimated_residents"] for b in residential_buildings)
        difference = census_population - assigned_population

        if difference > 0:
            for _ in range(difference):
                random.choice(residential_buildings)["estimated_residents"] += 1
        elif difference < 0:
            for _ in range(abs(difference)):
                random.choice(residential_buildings)["estimated_residents"] -= 1

        # Assign households
        households = []
        households += [("1_comp_work", profiles["1_comp_work"]) for _ in range(int(census["HH_1 comp"] * census_occupied))]
        households += [("1_comp_ret", profiles["1_comp_ret"]) for _ in range(int(census["HH_1 comp"] * (1 - census_occupied)))]
        households += [("2_comp_work", profiles["2_comp_work"]) for _ in range(int(census["HH_2 comp"] * census_occupied))]
        households += [("2_comp_ret", profiles["2_comp_ret"]) for _ in range(int(census["HH_2 comp"] * (1 - census_occupied)))]
        households += [("3_comp", profiles["3_comp"]) for _ in range(int(census["HH_3 comp"]))]
        households += [("4_comp_more", profiles["4_comp_more"]) for _ in range(int(census["HH_4 comp"] + census["HH_5 comp"] + census["HH_6 comp or more"]))]

        random.shuffle(households)

        for building in residential_buildings:
            building_id = str(building["ID"])
            num_households = max(1, round(building["estimated_residents"] / census_population * total_census_households))

            assigned_profiles = []
            for _ in range(num_households):
                if households:
                    ncomp_type, _ = households.pop(0)
                    assigned_profiles.append(ncomp_type)

            # ---- JSON SUMMARY ----
            summary = {
                "single_worker": {"Wasteful": 0, "Average": 0, "Saver": 0},
                "single_retired": {"Wasteful": 0, "Average": 0, "Saver": 0},
                "couple_workers": {"Wasteful": 0, "Average": 0, "Saver": 0},
                "couple_retired": {"Wasteful": 0, "Average": 0, "Saver": 0},
                "families": {"Wasteful": 0, "Average": 0, "Saver": 0}
            }

            for hh in assigned_profiles:
                if hh == "1_comp_work":
                    summary["single_worker"]["Average"] += 1
                elif hh == "1_comp_ret":
                    summary["single_retired"]["Average"] += 1
                elif hh == "2_comp_work":
                    summary["couple_workers"]["Average"] += 1
                elif hh == "2_comp_ret":
                    summary["couple_retired"]["Average"] += 1
                elif hh in ["3_comp", "4_comp_more"]:
                    summary["families"]["Average"] += 1

            building_summary[building_id] = summary
            building_estimated_residents[building_id] = building["estimated_residents"]

            # ---- Save per-building CSV ----
            if assigned_profiles:
                df = pd.DataFrame(assigned_profiles, columns=["Household Type"])
                file_path = os.path.join(output_folder, f"{building_id}.csv")
                df.to_csv(file_path, index=False)

    # ---- SAVE ONE JSON FILE AT THE END ----
    output_json_path = os.path.join(output_folder, "building_household_summary.json")
    with open(output_json_path, "w") as f:
        json.dump(building_summary, f, indent=2)
    print(f"✅ Saved household summary JSON to {output_json_path}")

    residents_json_path = os.path.join(output_folder, "building_estimated_residents.json")
    with open(residents_json_path, "w") as f:
        json.dump(building_estimated_residents, f, indent=2)


# Run function
assign_households_to_residential_buildings(census_data, building_data)
