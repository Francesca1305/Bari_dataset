import pandas as pd
import numpy as np
from matplotlib import pyplot as plt
import json
import random
import openpyxl
import os


census_data = "Data/census_data_out.json"
building_data = "Data/building_data_out.json"
profile_file = 'Data/profiles_HH.xlsx'
output_folder = "Data/Building_profiles"

#household_types = {"IMW":"isolated_members_workers",
#                   "IMR":"isolated_members_retired",
#                   "CWC":"couple_with_children",
#                   "CnoCW":"couple_without_children_workers",
#              "CnoCR":"couple_without_children_retired",
#              "SP":"single parents"}
#da rivedere se farlo con righe o con pagine dell'excel separate

ncomp_types = {"1component_worker":"1_comp_work",
               "1component_retired":"1_comp_ret",
               "2components_working":"2_comp_work",
               "2components_retired":"2_comp_ret",
               "3components":"3_comp",
               "4components_more":"4_comp_more"}

with open(census_data, 'r') as f:
    census_data = json.load(f)

with open(building_data, 'r') as f:
    building_data = json.load(f)

# read the xlsx based on the dictionary
profiles = {
    ncomp_types[sheet]: pd.read_excel(profile_file, sheet_name=sheet).values.tolist()
    for sheet in ncomp_types
}


def assign_HH_to_buildings(census_sections, building_data):
    assignments = []

     # variables from json
    for feature in census_sections["features"]:
        census = feature["properties"]
        census_id = census["SEZ21"]
        census_population = census["total resident population"]
        total_census_households = census["total households"]
        census_occupied = ((census["Italian occupied_IT10"] +
                            census["Foreign occupied_ST31"])/
                           census["total resident population"]) if census_population>0 else 0

        # filter residential buildings only
        residential_buildings = [b["properties"] for b in building_data["features"]
                                 if b["properties"]["SEZ21"] == census_id and b["properties"]["function"]
                                 in [11, 12]] # change with residential code
        if not residential_buildings:
            continue

        # Calculate total heated area for residential buildings only
        total_heated_area = sum(b["Shape_Area"] * b["nfloors"] for b in residential_buildings)
        avg_area_per_person = total_heated_area / census_population if census_population > 0 else 1

        # Calculate estimated number of residents per residential building
        for building in residential_buildings:
            building["heated_area"] = building["Shape_Area"] * building["nfloors"]
            building["estimated_residents"] = max(1, round(building["heated_area"]/
                                                           avg_area_per_person))

        # Ensure total assigned population match section population
        assigned_population = sum(b["estimated_residents"] for b in residential_buildings)
        difference = census_population - assigned_population

        # Adjust the number of residents per buildings by sum or subtraction
        if difference > 0:
            for _ in range(difference):
                random.choice(residential_buildings)["estimated_residents"] += 1
        elif difference < 0:
            for _ in range(abs(difference)):
                random.choice(residential_buildings)["estimated_residents"] -= 1

        # Ensure that at least one household is assigned to each building
        households = []
        for building in residential_buildings:
            building["estimated_residents"] = max(1, building["estimated_residents"])

        # List of family data based on census data. Selection of a random profile from xlsx file
        # for the range of household types
        # single working
        households += [("1_comp_ret", profiles["1_comp_work"]) for _ in
                       range(int(census["HH_1 comp"]*census_occupied))]
        # single retired
        households += [("1_comp_ret", profiles["1_comp_ret"]) for _ in
                       range(int(census["HH_1 comp"] * (1-census_occupied)))]
        # couple without children working
        households += [("2_comp_work", profiles["2_comp_work"]) for _ in
                       range(int(census["HH_2 comp"]*census_occupied))]
        # couple without children retired
        households += [("2_comp_ret", profiles["2_comp_ret"]) for _ in
                       range(int(census["HH_2 comp"] * (1-census_occupied)))]
        # couple with one child
        households += [("3_comp", profiles["3_comp"]) for _ in
                       range(int(census["HH_3 comp"]))]
        # couple with more than one child
        households += [("4_comp_more", profiles["4_comp_more"]) for _ in
                       range(int(census["HH_4 comp"] + census["HH_5 comp"]
                                 + census["HH_6 comp or more"]))]

        random.shuffle(households)

        # Initial assignment of households to buildings
        building_assignments = {b["ID"]: [] for b in residential_buildings}
        total_assigned_households = 0

        for building in residential_buildings:
            building["assigned_households"] = max(1, round(
                building["estimated_residents"] / census_population * total_census_households))

        # Correct number of assigned households
        assigned_households_check = sum(b["assigned_households"]
                                        for b in residential_buildings)
        diff_households = total_census_households - assigned_households_check

        while diff_households != 0:
            balance = random.choice(residential_buildings)
            if diff_households > 0:
                balance["assigned_households"] +=1
                diff_households -= 1
            elif diff_households < 0 and balance["assigned_households"] > 1:
                balance["assigned_households"] -=1
                diff_households +=1

        # Final assignment of households
        for building in residential_buildings:
            for _ in range(building["assigned_households"]):
                if households:
                    ncomp_type, profile = households.pop(0)
                    building_assignments[building["ID"]].append((ncomp_type, profile))
                    total_assigned_households += 1


        for building in residential_buildings:
            building_ID = building["ID"]
            assignments.append({
                "building_ID": building_ID,
                "assignments": building_assignments[building_ID]
            })

        # Create a DataFrame for the building assignments

        print(building_assignments["116"])
        for building in residential_buildings:
            building_id = building["ID"]
            print(building_id)
            # Convert the assignment data for the building into a DataFrame
            #building_data = building_assignments[building_id]
            #building_df = pd.DataFrame(building_data, columns=["Household Type", "Profile"])

            # Save the DataFrame to a CSV file
            #file_path = os.path.join(output_folder, f"{building_id}.csv")
            #building_df.to_csv(file_path, index=False)

    return assignments


# Call the function
assignments = assign_HH_to_buildings(census_data, building_data)






