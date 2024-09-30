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
    r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Excel\Formato Infraestructura - Perfiles y Capacidades - Ajustado.xlsx', sheet_name='Laboratorios', dtype=object)
dataframe1 = dataframe1.fillna('')
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
    temp_lab_type = temp_lab["tipo"].values[0]
    temp_lab_parent_type= temp_lab["tipo-parent"].values[0]

    # Detect empty fields and replacing with empty text
    if (isinstance(temp_lab_address, int) | isinstance(temp_lab_address, float)):
        temp_project_description = ""
    else:
        temp_project_description = temp_lab_address.replace("\n", " ")

    if (isinstance(temp_lab_url, str) and validators.url(temp_lab_url)):
        temp_lab_url = temp_lab_url
    else:
        temp_lab_url = ""


    if isinstance(temp_lab_created, datetime.datetime):
        temp_lab_created = temp_lab_created.strftime("%Y-%m-%d")

    # línea para manejar numpy.datetime64
    if isinstance(temp_lab_created, np.datetime64):
        temp_lab_created = str(temp_lab_created.astype('datetime64[D]'))

    equipment = ET.SubElement(items, "equipment", externalId=temp_lab_id)

    title = ET.SubElement(equipment, "title", formatted="true")
    ET.SubElement(title, "text", locale="es_CO").text = temp_lab_es_name
    ET.SubElement(title, "text", locale="en_US").text = temp_lab_eng_name

    type = ET.SubElement(equipment, "type", uri="/dk/atira/pure/equipment/equipmenttypes/equipment/" + temp_lab_type)
    type_term = ET.SubElement(type, "term", formatted = "false")
    if temp_lab_type == "laboratory":
        ET.SubElement(type_term, "text", locale="en_US").text = "Laboratory"
        ET.SubElement(type_term, "text", locale="es_CO").text = "Laboratorio"
    elif temp_lab_type == "area":
        ET.SubElement(type_term, "text", locale="en_US").text = "Area"
        ET.SubElement(type_term, "text", locale="es_CO").text = "Área"
    elif temp_lab_type == "academic_unit":
        ET.SubElement(type_term, "text", locale="en_US").text = "Academic Unit"
        ET.SubElement(type_term, "text", locale="es_CO").text = "Unidad Académica"

    category = ET.SubElement(equipment, "category", uri="/dk/atira/pure/equipment/category")
    
    # Descriptions
    if temp_lab_es_description != "" and temp_lab_eng_description != "":
        descriptions = ET.SubElement(equipment, "descriptions")
        description = ET.SubElement(descriptions, "description")
        value = ET.SubElement(description, "value", formatted="false")
        ET.SubElement(value, "text", locale="es_CO").text = temp_lab_es_description
        ET.SubElement(value, "text", locale="en_US").text = temp_lab_eng_description
        description_type = ET.SubElement(description, "type", uri="/dk/atira/pure/equipment/descriptions/equipmentdescription")
        description_type_term = ET.SubElement(description_type, "term", formatted="false")
        ET.SubElement(description_type_term, "text", locale="es_CO").text = "Descripción"
        ET.SubElement(description_type_term, "text", locale="en_US").text = "Description"


    # Equipment details
    equipment_details = ET.SubElement(equipment, "equipmentDetails")
    equipment_detail = ET.SubElement(equipment_details, "equipmentDetail")
    name_detail = ET.SubElement(equipment_detail, "name", formatted="true")
    ET.SubElement(name_detail, "text",locale="en_US").text = temp_lab_es_name
    ET.SubElement(name_detail, "text",locale="es_CO").text = temp_lab_eng_name

    ids_detail = ET.SubElement(equipment_detail, "ids")
    ids_detail_id = ET.SubElement(ids_detail, "id")
    ET.SubElement(ids_detail_id, "value", formatted="false").text = temp_lab_id
    type_detail = ET.SubElement(ids_detail_id, "type", uri="/dk/atira/pure/equipment/equipmentsources/internalid")
    type_detail_term = ET.SubElement(type_detail, "term", formatted="false")
    ET.SubElement(type_detail_term, "text", locale="es_CO").text = "ID interna"
    ET.SubElement(type_detail_term, "text", locale="en_US").text = "Internal ID"
    ET.SubElement(equipment_detail, "acquisitionDate").text = temp_lab_created
    # equipment_details_ids = ET.SubElement(equipment_detail, "ids")
    # equipment_details_ids_id = ET.SubElement(equipment_details_ids, 


    person_associations = ET.SubElement(equipment, "personAssociations")
    person_association = ET.SubElement(
        person_associations, "personAssociation")
    ET.SubElement(person_association, "person", externalId=temp_lab_person_id)
    ET.SubElement(person_association, "personRole",
                  uri="/dk/atira/pure/equipment/roles/equipment/manager")
    person_org_units = ET.SubElement(person_association, "organisationalUnits")
    # TODO revisar que valor va como externalId
    person_org_unit = ET.SubElement(
        person_org_units, "organisationalUnit", externalId="218")

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

    # Contact Persons
    contact_persons = ET.SubElement(equipment, "contactPersons")
    contact_person = ET.SubElement(
        contact_persons, "contactPerson", externalId=temp_lab_person_id, externalIdSource="synchronisedPerson",externallyManaged="true")
    name_contact = ET.SubElement(contact_person, "name", formatted="false")
    ET.SubElement(name_contact, "text").text = temp_lab_person

    # Addresses
    addresses = ET.SubElement(equipment, "addresses")
    address = ET.SubElement(addresses, "address")
    address_type = ET.SubElement(address, "addressType", uri="/dk/atira/pure/equipment/equipmentaddresstype/postal")
    # address_type_term = ET.SubElement(address, "term", formatted="false")
    # ET.SubElement(address_type_term, "text", locale="es_CO").text = "Dirección postal"
    # ET.SubElement(address_type_term, "text", locale="en_US").text = "Postal address"
    ET.SubElement(address, "street").text = temp_lab_address
    # Placeholder
    ET.SubElement(address, "building").text = "Edificio José Gabriel Maldonado S.J.-Laboratorios"
    ET.SubElement(address, "city").text = "Bogotá D.C."  # Placeholder
    country = ET.SubElement(address, "country", uri="/dk/atira/pure/core/countries/co")
    country_term = ET.SubElement(country, "term", formatted="false")
    ET.SubElement(country_term, "text", locale="es_CO").text = "Colombia"
    ET.SubElement(country_term, "text", locale="en_US").text = "Colombia"

    # Phone Numbers
    phone_numbers = ET.SubElement(equipment, "phoneNumbers")
    phone_number = ET.SubElement(phone_numbers, "phoneNumber")
    ET.SubElement(phone_number, "value", formatted="true").text = str(temp_lab_phone)
    phone_number_type = ET.SubElement(phone_number, "type", uri="/dk/atira/pure/equipment/equipmentphonenumbertype/phone")
    phone_number_type_term = ET.SubElement(phone_number_type, "term", formatted="false")
    ET.SubElement(phone_number_type_term, "text", locale="es_CO").text = "Teléfono"
    ET.SubElement(phone_number_type_term, "text", locale="en_US").text = "Phone"

    # Web Addresses
    if temp_lab_url != "":
        web_addresses = ET.SubElement(equipment, "webAddresses")
        web_address = ET.SubElement(web_addresses, "webAddress")
        web_address_value = ET.SubElement(
            web_address, "value", formatted="false")
        ET.SubElement(web_address_value, "text", locale="es_CO").text = temp_lab_url
        ET.SubElement(web_address_value, "text", locale="en_US").text = temp_lab_url
        web_type = ET.SubElement(web_address, "type", uri="/dk/atira/pure/equipment/equipmentwebaddresstype/website")
        web_type_term = ET.SubElement(web_type, "term", formatted="false") 
        ET.SubElement(web_type_term, "text", locale="es_CO").text = "Sitio web"
        ET.SubElement(web_type_term, "text", locale="en_US").text = "Website"

    # Loan Type
    # set to lowercase and trim any space
    temp_lab_loan_type = temp_lab_loan_type.lower().strip()
    loan_type_spanish = ""
    loan_type_english = ""
    if temp_lab_loan_type == "ambos":
        temp_lab_loan_type = "/dk/atira/pure/equipment/loantypes/internalexternal"
        loan_type_spanish = "Disponible para el préstamo, a nivel interno y externo"
        loan_type_english = "Available for loan - internal and external"
    elif temp_lab_loan_type == "interno":
        temp_lab_loan_type = "/dk/atira/pure/equipment/loantypes/internal"
        loan_type_spanish = "Disponible para el préstamo, solo a nivel interno"
        loan_type_english = "Available for loan - internal only"
    else:
        temp_lab_loan_type = "/dk/atira/pure/equipment/loantypes/notavailable"
        loan_type_spanish = "No está disponible para el préstamo"
        loan_type_english = "Not available for loan"

    # Loan terms and certifications
    if temp_lab_certifications != "No aplica":
        # concatenamos las certificaciones a los servicios
        temp_lab_services = temp_lab_services + " " + temp_lab_certifications
    
    if temp_lab_services == "":
        temp_lab_services = "No aplica"

    loan_type = ET.SubElement(equipment, "loanType", uri=temp_lab_loan_type)
    loan_type_term = ET.SubElement(loan_type, "term", formatted="false")
    ET.SubElement(loan_type_term, "text", locale="en_US").text = loan_type_english
    ET.SubElement(loan_type_term, "text", locale="es_CO").text = loan_type_spanish
    loan_term = ET.SubElement(equipment, "loanTerm", formatted="true")
    loan_term_text = ET.SubElement(loan_term, "text", locale="es_CO")
    loan_term_text.text = temp_lab_services

    # parents
    if temp_lab_parent_type != "":
        parents_group = ET.SubElement(equipment, "parents")
        parent = ET.SubElement(parents_group, "parent",externalId=temp_lab_area_id)
        parent_name = ET.SubElement(parent, "name", formatted="false")
        ET.SubElement(parent_name, "text", locale="es_CO").text = temp_lab_area
        parent_type = ET.SubElement(parent, "type", uri="/dk/atira/pure/equipment/equipmenttypes/equipment/" + temp_lab_parent_type)
        parent_type_term = ET.SubElement(parent_type, "term", formatted="false")
        if parent_type_term == "area":
            ET.SubElement(type_term, "text", locale="en_US").text = "Area"
            ET.SubElement(type_term, "text", locale="es_CO").text = "Área"
        elif parent_type_term == "academic_unit":
            ET.SubElement(type_term, "text", locale="en_US").text = "Academic Unit"
            ET.SubElement(type_term, "text", locale="es_CO").text = "Unidad Académica"

    # Keywords
    if temp_lab_keywords != "":
        keyword_groups = ET.SubElement(equipment, "keywordGroups")
        keyword_group = ET.SubElement(keyword_groups, "keywordGroup", logicalName="keywordContainers")
        keyword_group_type = ET.SubElement(keyword_group, "type")
        keyword_group_type_term = ET.SubElement(keyword_group_type, "term", formatted="false")
        ET.SubElement(keyword_group_type_term, "text", locale="es_CO").text = "Palabras clave"
        ET.SubElement(keyword_group_type_term, "text", locale="en_US").text = "Keywords"
        keyword_containers = ET.SubElement(keyword_group, "keywordContainers")
        keyword_container = ET.SubElement(keyword_containers, "keywordContainer")
        free_keywords = ET.SubElement(keyword_container, "freeKeywords")
        free_keyword = ET.SubElement(free_keywords, "freeKeyword", locale="es_CO")
        free_keywords_list = ET.SubElement(free_keyword, "freeKeywords")
        # Split keywords and add them to the XML
        for keyword in temp_lab_keywords.split(','):
            free_keyword = ET.SubElement(free_keywords_list, "freeKeyword")
            free_keyword.text = keyword.strip()


# Convert the ElementTree to a string
tree = ET.ElementTree(result)
xml_str = ET.tostring(result, encoding="unicode", method="xml")

# Save to a file
with open(r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Resultado\2024_09_30_LABS_KV_V10.xml', "w", encoding="utf-8") as f:
    # Escribir la declaración manualmente
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    # Escribir el contenido del árbol XML
    f.write(xml_str)

print("xml generado")
