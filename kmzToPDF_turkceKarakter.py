import os
import zipfile
import xml.etree.ElementTree as ET
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter

def kmz_to_kml(kmz_path, output_kml_path):
    with zipfile.ZipFile(kmz_path, 'r') as kmz:
        for file in kmz.namelist():
            if file.endswith('.kml'):
                with open(output_kml_path, 'wb') as kml_file:
                    kml_file.write(kmz.read(file))
                print(f"KML file extracted to {output_kml_path}")
                return
    raise FileNotFoundError("No KML file found in the KMZ archive.")

def parse_kml_to_pdf(kml_path, pdf_path):
    tree = ET.parse(kml_path)
    root = tree.getroot()

    # Namespaces may be required to parse the KML correctly
    namespaces = {'kml': 'http://www.opengis.net/kml/2.2'}

    # Find all Placemark elements
    placemarks = root.findall('.//kml:Placemark', namespaces)

    styles = getSampleStyleSheet()
    styles['Normal'].fontName = 'Helvetica'
    styles['Title'].fontName = 'Helvetica-Bold'

    story = []

    story.append(Paragraph("KML Data Export", styles['Title']))
    story.append(Spacer(1, 12))

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
                coordinates_link = f'<a href="https://www.google.com/maps?q={lat},{lon}">View on Google Maps</a>'
            else:
                coordinates_link = 'Invalid coordinates'
        else:
            coordinates_link = 'N/A'

        story.append(Paragraph(f"<b>Name:</b> {name_text}", styles['Normal']))
        story.append(Paragraph(f"<b>Description:</b> {description_text}", styles['Normal']))
        story.append(Paragraph(f"<b>Coordinates:</b> {coordinates_link}", styles['Normal']))
        story.append(Spacer(1, 12))

    doc = SimpleDocTemplate(pdf_path, pagesize=letter)
    doc.build(story)
    print(f"PDF file saved to {pdf_path}")

if __name__ == "__main__":
    kmz_path = "./fatih.kmz"
    kml_path = "./fatih.kml"
    pdf_path = "./fatih.pdf"

    # Convert KMZ to KML
    kmz_to_kml(kmz_path, kml_path)

    # Parse KML and write to PDF
    parse_kml_to_pdf(kml_path, pdf_path)