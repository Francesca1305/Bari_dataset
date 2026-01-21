import json
from shapely.geometry import shape, mapping, MultiPolygon, Polygon
from shapely.ops import unary_union
from pathlib import Path

def enforce_right_hand_rule(geojson):
    for feature in geojson['features']:
        geometry = shape(feature['geometry'])
        if isinstance(geometry, Polygon):
            if not geometry.exterior.is_ccw:
                geometry = Polygon(geometry.exterior.coords[::-1], [interior.coords[::-1] for interior in geometry.interiors])
        elif isinstance(geometry, MultiPolygon):
            new_polygons = []
            for polygon in geometry.geoms:  # Use .geoms to iterate over the individual polygons
                if not polygon.exterior.is_ccw:
                    polygon = Polygon(polygon.exterior.coords[::-1], [interior.coords[::-1] for interior in polygon.interiors])
                new_polygons.append(polygon)
            geometry = MultiPolygon(new_polygons)
        feature['geometry'] = mapping(geometry)
    return geojson

data_path = Path(__file__).parent.parent / 'hub/data'  # Adjust this path as needed
# Load the GeoJSON file
input_filepath = data_path / 'C:/Users/frenc/Politecnico Di Torino Studenti Dropbox/Francesca Vecchi/PhD/Conferences/Bari CEES 2025/Random assignation/Data/census_data.json'
output_filepath = data_path / 'C:/Users/frenc/Politecnico Di Torino Studenti Dropbox/Francesca Vecchi/PhD/Conferences/Bari CEES 2025/Random assignation/Data/census_data_out.json'

with open(input_filepath, 'r') as f:
    geojson_data = json.load(f)

# Iterate through the features and convert MultiPolygons to Polygons
for feature in geojson_data['features']:
    geom = shape(feature['geometry'])
    if isinstance(geom, MultiPolygon):
        # Merge the polygons into a single polygon
        merged_polygon = unary_union(geom)
        # Convert the merged polygon back to GeoJSON format
        feature['geometry'] = mapping(merged_polygon)

# Enforce the right-hand rule
geojson_data = enforce_right_hand_rule(geojson_data)

# Save the updated GeoJSON file
with open(output_filepath, 'w') as f:
    json.dump(geojson_data, f)
