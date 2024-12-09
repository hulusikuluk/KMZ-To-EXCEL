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

        name_text = name.text if name is not None else 'N/A'
        description_text = description.text if description is not None else 'N/A'
        coordinates_text = coordinates.text.strip() if coordinates is not None else ''

        if coordinates_text:
            coord_parts = coordinates_text.split(',')
            if len(coord_parts) >= 2:
                lon, lat = coord_parts[0], coord_parts[1]
                coordinates_link = f"https://earth.google.com/web/@{lat},{lon},10000a,100d"
            else:
                coordinates_link = 'Invalid coordinates'
        else:
            coordinates_link = 'N/A'

        data.append({
            'Name': name_text,
            'Description': description_text,
            'Coordinates (Google Earth Link)': coordinates_link
        })

    # Create a DataFrame and adjust rows for multi-line descriptions
    df = pd.DataFrame(data)

    # Write to Excel
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='KML Data')

        # Adjust column widths and enable text wrapping for the Description column
        workbook = writer.book
        worksheet = writer.sheets['KML Data']
        worksheet.set_column('A:A', 20)  # Name column
        worksheet.set_column('B:B', 50, workbook.add_format({'text_wrap': True}))  # Description column
        worksheet.set_column('C:C', 50)  # Coordinates (Google Earth Link) column

    print(f"Excel file saved to {excel_path}")

if __name__ == "__main__":
    kmz_path = "./fatih.kmz"
    kml_path = "./fatih.kml"
    excel_path = "./fatih.xlsx"

    # Convert KMZ to KML
    kmz_to_kml(kmz_path, kml_path)

    # Parse KML and write to Excel
    parse_kml_to_excel(kml_path, excel_path)
