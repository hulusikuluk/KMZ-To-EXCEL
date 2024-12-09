import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd

def kmz_to_kml(kmz_path, output_kml_path):
    with zipfile.ZipFile(kmz_path, 'r') as kmz:
        for file in kmz.namelist():
            if file.endswith('.kml'):
                with open(output_kml_path, 'wb') as kml_file:
                    kml_file.write(kmz.read(file))
                print(f"KML file extracted to {output_kml_path}")
                return
    raise FileNotFoundError("No KML file found in the KMZ archive.")

def parse_kml_to_excel(kml_path, excel_path):
    tree = ET.parse(kml_path)
    root = tree.getroot()

    # Namespaces may be required to parse the KML correctly
    namespaces = {'kml': 'http://www.opengis.net/kml/2.2'}

    # Find all Placemark elements
    placemarks = root.findall('.//kml:Placemark', namespaces)

    data = []

    for placemark in placemarks:
        name = placemark.find('kml:name', namespaces)
        description = placemark.find('kml:description', namespaces)
        coordinates = placemark.find('.//kml:coordinates', namespaces)

        data.append({
            'Name': name.text if name is not None else '',
            'Description': description.text if description is not None else '',
            'Coordinates': coordinates.text.strip() if coordinates is not None else ''
        })

    # Create a DataFrame and save to Excel
    df = pd.DataFrame(data)
    df.to_excel(excel_path, index=False)
    print(f"Excel file saved to {excel_path}")
if __name__ == "__main__":
    kmz_path = "./fatih.kmz"
    kml_path = "./fatih.kml"
    excel_path = "./fatih.xlsx"

    # Convert KMZ to KML
    kmz_to_kml(kmz_path, kml_path)

    # Parse KML and write to Excel
    parse_kml_to_excel(kml_path, excel_path)
