import json
building_data_path = './Data/building_data_out.json'
census_data_path = './Data/census_data_out.json'
with open(building_data_path, 'r') as file:
    building_data = json.load(file)

with open(census_data_path, 'r') as file1:
    census_data = json.load(file1)
for feature in census_data["features"]:
    census = feature["properties"]
    census_id = census["SEZ21"]
    residential_buildings = [b["properties"] for b in building_data["features"]
                                 if b["properties"]["SEZ21"] == census_id and b["properties"]["function"]
                                 in [11, 12]]

