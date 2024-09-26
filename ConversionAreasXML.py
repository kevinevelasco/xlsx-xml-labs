from urllib.parse import urlparse
import pandas as pd
from spacy.language import Language
from spacy_langdetect import LanguageDetector
import xml.etree.ElementTree as ET
import validators
import datetime
import numpy as np

# Create Element Tree root class
result = ET.Element("result")
result.set('xmlns:xsi', "http://www.w3.org/2001/XMLSchema-instance")
result.set('xsi:schemaLocation',
           "http://localhost:8080/ws/api/524/xsd/schema1.xsd")
items = ET.SubElement(result, "items")

# Read excel and convert to dataframe
dataframe1 = pd.read_excel(
    r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Excel\Formato Infraestructura - Perfiles y Capacidades - Ajustado.xlsx', sheet_name='Areas', dtype=object)
dataframe1 = dataframe1.fillna('')
column = dataframe1.id_unico.unique()

for uid in column:
    # Create Element Tree root class
    temp_area = dataframe1[dataframe1["id_unico"] == uid]
    temp_area_id = str(uid)
    temp_area_name = temp_area["Nombre español"].values[0]

    equipment = ET.SubElement(items, "equipment", externalId=temp_area_id)

    title = ET.SubElement(equipment, "title", formatted="true")
    ET.SubElement(title, "text", locale="es_CO").text = temp_area_name

    type = ET.SubElement(
        equipment, "type", uri="/dk/atira/pure/equipment/equipmenttypes/equipment/area")
    type_term = ET.SubElement(type, "term", formatted = "false")
    ET.SubElement(type_term, "text", locale="en_US").text = "Area"
    ET.SubElement(type_term, "text", locale="es_CO").text = "Área"

    category = ET.SubElement(
        equipment, "category", uri="/dk/atira/pure/equipment/category")

    org_units = ET.SubElement(equipment, "organisationalUnits")
    org_unit = ET.SubElement(org_units, "organisationalUnit", externalId="218")
    name = ET.SubElement(org_unit, "name", formatted="false")
    # TODO placeholder
    ET.SubElement(name, "text", locale="es_CO").text = "Facultad de Ingeniería"
    # TODO placeholder
    ET.SubElement(name, "text", locale="en_US").text = "School of Engineering"


    # Managing Organizational Unit
    managing_org_unit = ET.SubElement(
        equipment, "managingOrganisationalUnit", externalId="218", externallyManaged="true")
    name_org = ET.SubElement(managing_org_unit, "name", formatted="false")
    ET.SubElement(name_org, "text",
                  locale="es_CO").text = "Facultad de Ingeniería"
    ET.SubElement(name_org, "text",
                  locale="en_US").text = "School of Engineering"
    type_org = ET.SubElement(managing_org_unit, "type", uri="/dk/atira/pure/organisation/organisationtypes/organisation/faculty")
    type_org_term = ET.SubElement(type_org, "term", formatted="false")
    ET.SubElement(type_org_term, "text", locale="es_CO").text = "Facultad"
    ET.SubElement(type_org_term, "text", locale="en_US").text = "Faculty"



# Convert the ElementTree to a string
tree = ET.ElementTree(result)
xml_str = ET.tostring(result, encoding="unicode", method="xml")

# Save to a file
with open(r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Resultado\2024_09_26_AREAS_KV_V1.xml', "w", encoding="utf-8") as f:
    # Escribir la declaración manualmente
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    # Escribir el contenido del árbol XML
    f.write(xml_str)

print("xml generado")
