import pandas as pd
from spacy.language import Language
from spacy_langdetect import LanguageDetector
import xml.etree.ElementTree as ET
import datetime

# Create Element Tree root class
result = ET.Element("result")
result.set('xmlns:xsi', "http://www.w3.org/2001/XMLSchema-instance")
result.set('xsi:schemaLocation',
           "http://localhost:8080/ws/api/524/xsd/schema1.xsd")
items = ET.SubElement(result, "items")

# Read excel and convert to dataframe
dataframe1 = pd.read_excel(
    r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Excel\Formato Infraestructura - Perfiles y Capacidades.xlsx', sheet_name='Laboratorios', dtype=object)

column = dataframe1.id_unico.unique()

for uid in column:
    # Create Element Tree root class
    temp_lab = dataframe1[dataframe1["id_unico"] == uid]
    temp_lab_id = str(uid)
    temp_lab_area = temp_lab["Area"].values[0]
    temp_lab_area_id = temp_lab["area-id"].values[0]
    temp_lab_eng_name = temp_lab["Nombre Inglés"].values[0]
    temp_lab_es_name = temp_lab["Nombre español"].values[0]
    temp_lab_eng_description = temp_lab["Descripción Inglés"].values[0]
    temp_lab_es_description = temp_lab["Descripción Español"].values[0]
    temp_lab_address = temp_lab["Dirección"].values[0]
    temp_lab_phone = temp_lab["Número de contacto"].values[0]
    temp_lab_url = temp_lab["URL"].values[0]
    temp_lab_person = temp_lab["Nombre encargado"].values[0]
    temp_lab_person_id = temp_lab["ID Empleado"].values[0]
    temp_lab_created = temp_lab["Fecha creación Laboratorios"].values[0]
    temp_lab_unidad = "Departamento de Ingeniería de Sistemas"
    temp_lab_loan_type = temp_lab["Disponible para préstamo"].values[0]
    temp_lab_keywords = temp_lab["Palabras clave"].values[0]
    temp_lab_services = temp_lab["Servicios del Laboratorio"].values[0]
    temp_lab_certifications = temp_lab["Certificaciones"].values[0]

    # Detect empty fields and replacing with empty text
    if (isinstance(temp_lab_address, int) | isinstance(temp_lab_address, float)):
        temp_project_description = ""
    else:
        temp_project_description = temp_lab_address.replace("\n", " ")

    if (isinstance(temp_lab_url, str)):
        temp_lab_url = temp_lab_url
    else:
        temp_lab_url = ""

    if isinstance(temp_lab_created, datetime.datetime):
        temp_lab_created = temp_lab_created.strftime("%Y-%m-%d")

    equipment = ET.SubElement(items, "equipment", externalId=temp_lab_id)

    title = ET.SubElement(equipment, "title", formatted="true")
    ET.SubElement(title, "text", locale="es_CO").text = temp_lab_es_name
    ET.SubElement(title, "text", locale="en_US").text = temp_lab_eng_name

    type = ET.SubElement(
        equipment, "type", uri="/dk/atira/pure/equipment/equipmenttypes/laboratory")
    category = ET.SubElement(
        equipment, "category", uri="/dk/atira/pure/equipment/category/classification")

    person_associations = ET.SubElement(equipment, "personAssociations")
    person_association = ET.SubElement(person_associations, "personAssociation")
    ET.SubElement(person_association, "person", externalId=temp_lab_person_id)
    ET.SubElement(person_association, "personRole", uri="/dk/atira/pure/equipment/roles/equipment/manager")
    person_org_units = ET.SubElement(person_association, "organisationalUnits")
    person_org_unit = ET.SubElement(person_org_units, "organisationalUnit", externalId="218") #TODO revisar que valor va como externalId
    

    org_units = ET.SubElement(equipment, "organisationalUnits")
    org_unit = ET.SubElement(org_units, "organisationalUnit", externalId="218")
    name = ET.SubElement(org_unit, "name", formatted="false")
    # TODO placeholder
    ET.SubElement(name, "text", locale="es_CO").text = "Facultad de Ingeniería"
    # TODO placeholder
    ET.SubElement(name, "text", locale="en_US").text = "Facultad de Ingeniería"

    # Descriptions
    descriptions = ET.SubElement(equipment, "descriptions")
    description = ET.SubElement(descriptions, "description")
    value = ET.SubElement(description, "value", formatted="false")
    ET.SubElement(value, "text", locale="es_CO").text = temp_lab_es_description
    ET.SubElement(
        value, "text", locale="en_US").text = temp_lab_eng_description
    ET.SubElement(description, "type",
                  uri="/dk/atira/pure/equipment/descriptions/equipmentdescription")

    # Equipment details
    equipment_details = ET.SubElement(equipment, "equipmentDetails")
    equipment_detail = ET.SubElement(equipment_details, "equipmentDetail")
    name_detail = ET.SubElement(equipment_detail, "name", formatted="true")
    ET.SubElement(name_detail, "text",
                  locale="es_CO").text = "Detalle en español"
    ET.SubElement(name_detail, "text",
                  locale="en_US").text = "Detail in English"
    ET.SubElement(equipment_detail, "acquisitionDate").text = temp_lab_created

    # Managing Organizational Unit
    managing_org_unit = ET.SubElement(
        equipment, "managingOrganisationalUnit", externalId="218")
    name_org = ET.SubElement(managing_org_unit, "name", formatted="false")
    ET.SubElement(name_org, "text",
                  locale="es_CO").text = "Facultad de Ingeniería"
    ET.SubElement(name_org, "text",
                  locale="en_US").text = "Facultad de Ingeniería"

    # Contact Persons
    contact_persons = ET.SubElement(equipment, "contactPersons")
    contact_person = ET.SubElement(
        contact_persons, "contactPerson", externalId=temp_lab_person_id)
    name_contact = ET.SubElement(contact_person, "name", formatted="false")
    ET.SubElement(name_contact, "text").text = temp_lab_person

    # Addresses
    addresses = ET.SubElement(equipment, "addresses")
    address = ET.SubElement(addresses, "address")
    ET.SubElement(address, "addressType",
                  uri="/dk/atira/pure/equipment/equipmentaddresstype/postal")
    ET.SubElement(address, "street").text = temp_lab_address
    # Placeholder
    ET.SubElement(
        address, "building").text = "Edificio José Gabriel Maldonado S.J.-Laboratorios"
    ET.SubElement(address, "city").text = "Bogotá D.C."  # Placeholder
    ET.SubElement(address, "country", uri="/dk/atira/pure/core/countries/co")

    # Phone Numbers
    phone_numbers = ET.SubElement(equipment, "phoneNumbers")
    phone_number = ET.SubElement(phone_numbers, "phoneNumber")
    ET.SubElement(phone_number, "value",
                  formatted="true").text = str(temp_lab_phone)
    ET.SubElement(phone_number, "type",
                  uri="/dk/atira/pure/equipment/equipmentphonenumbertype/phone")

    # Web Addresses
    if temp_lab_url != "":
        web_addresses = ET.SubElement(equipment, "webAddresses")
        web_address = ET.SubElement(web_addresses, "webAddress")
        web_address_value = ET.SubElement(
            web_address, "value", formatted="false")
        ET.SubElement(web_address_value, "text",
                      locale="es_CO").text = temp_lab_url
        ET.SubElement(web_address_value, "text",
                      locale="en_US").text = temp_lab_url
        ET.SubElement(web_address, "type",
                      uri="/dk/atira/pure/equipment/equipmentwebaddresstype/website")

    # Loan Type
    # set to lowercase and trim any space

    temp_lab_loan_type = temp_lab_loan_type.lower().strip()
    if temp_lab_loan_type == "ambos":
        temp_lab_loan_type = "/dk/atira/pure/equipment/loantypes/internalexternal"
    elif temp_lab_loan_type == "interno":
        temp_lab_loan_type = "/dk/atira/pure/equipment/loantypes/internal"
    else:
        temp_lab_loan_type = "/dk/atira/pure/equipment/loantypes/notavailable"

    # Loan terms and certifications
    if temp_lab_certifications != "No aplica":
        # concatenamos las certificaciones a los servicios
        temp_lab_services = temp_lab_services + " " + temp_lab_certifications

    loan_type = ET.SubElement(equipment, "loanType", uri=temp_lab_loan_type)
    loan_term = ET.SubElement(equipment, "loanTerm")
    ET.SubElement(loan_term, "text", locale="es_CO").text = temp_lab_services

    # parents
    parents_group = ET.SubElement(equipment, "parents")
    parent = ET.SubElement(parents_group, "parent",
                           externalId=temp_lab_area_id)
    parent_name = ET.SubElement(parent, "name", formatted="false")
    ET.SubElement(parent_name, "text", locale="es_CO").text = temp_lab_area

    # Keywords
    keyword_groups = ET.SubElement(equipment, "keywordGroups")
    keyword_group = ET.SubElement(
        keyword_groups, "keywordGroup", logicalName="keywordContainers")
    keyword_containers = ET.SubElement(keyword_group, "keywordContainers")
    keyword_container = ET.SubElement(keyword_containers, "keywordContainer")
    free_keywords = ET.SubElement(keyword_container, "freeKeywords")

    # Split keywords and add them to the XML
    for keyword in temp_lab_keywords.split(','):
        free_keyword = ET.SubElement(
            free_keywords, "freeKeyword", locale="es_CO")
        free_keyword.text = keyword.strip()

# Convert the ElementTree to a string
tree = ET.ElementTree(result)
xml_str = ET.tostring(result, encoding="unicode", method="xml")

# Save to a file
with open(r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Resultado\2024_09_17_LABS_KV_V6.xml', "w", encoding="utf-8") as f:
    # Escribir la declaración manualmente
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    # Escribir el contenido del árbol XML
    f.write(xml_str)

print("xml generado")
