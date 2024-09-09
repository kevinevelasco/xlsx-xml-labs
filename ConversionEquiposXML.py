import pandas as pd
from spacy.language import Language
from spacy_langdetect import LanguageDetector
import xml.etree.ElementTree as ET
import datetime

#Create Element Tree root class
result = ET.Element("result")
result.set('xmlns:xsi',"http://www.w3.org/2001/XMLSchema-instance")
result.set('xsi:schemaLocation',"http://localhost:8080/ws/api/524/xsd/schema1.xsd")
items = ET.SubElement(result,"items")

#Read excel and convert to dataframe
dataframe1 = pd.read_excel(r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Excel\Formato Infraestructura - Perfiles y Capacidades.xlsx', sheet_name='Equipos', dtype=object)

column = dataframe1.id_ubicacion.unique()

for uid in column:
    #Create Element Tree root class
    temp_eq = dataframe1[dataframe1["id_ubicacion"]==uid]
    temp_eq_id = str(uid)
    temp_eq_id_lab = temp_eq["id-lab"].values[0]
    temp_eq_lab_es_name = temp_eq["Espacio ubicación"].values[0]
    temp_eq_plaqueta = temp_eq["ID Plaqueta (Activos Fijos)"].values[0]
    temp_eq_es_name = temp_eq["Nombre equipo español"].values[0]
    temp_eq_eng_name = temp_eq["Nombre equipo ingles"].values[0]
    temp_eq_es_description = temp_eq["Descripción español"].values[0]
    temp_eq_eng_description = temp_eq["Descripción ingles"].values[0]
    temp_eq_adquisicion = temp_eq["Fecha de adquisición"].values[0]
    temp_eq_baja = temp_eq["Fecha estimada baja (Vida útil)"].values[0]
    temp_eq_siu = temp_eq["Id SIU responsable"].values[0]
    temp_eq_model = temp_eq["Modelo"].values[0]
    temp_eq_fabricante = temp_eq["Nombre fabricante"].values[0]
    temp_eq_pais_fabricante = temp_eq["Pais fabricante"].values[0]
    temp_eq_loan_type = temp_eq["Disponible para prestamo (interno/externo/ambos)"].values[0]
    temp_eq_keywords = temp_eq["Palabras clave"].values[0]
    temp_eq_services = temp_eq["Servicios"].values[0]

    if (isinstance(temp_eq_es_description,str)):
        temp_eq_es_description = temp_eq_es_description
    else:
        temp_eq_es_description = ""
    
    if (isinstance(temp_eq_eng_description,str)):
        temp_eq_eng_description = temp_eq_eng_description
    else:
        temp_eq_eng_description = ""

    if isinstance(temp_eq_adquisicion, datetime.datetime):
        temp_eq_adquisicion = temp_eq_adquisicion.strftime("%Y-%m-%d")

    if isinstance(temp_eq_baja, datetime.datetime):
        temp_eq_baja = temp_eq_baja.strftime("%Y-%m-%d")


    if isinstance(temp_eq_model, str):
        temp_eq_model = temp_eq_model
    else:
        temp_eq_model = ""

    equipment= ET.SubElement(items,"equipment", externalId=temp_eq_id)

    title = ET.SubElement(equipment, "title", formatted="true")
    ET.SubElement(title, "text", locale="es_CO").text = temp_eq_es_name
    ET.SubElement(title, "text", locale="en_US").text = temp_eq_eng_name

    type = ET.SubElement(equipment, "type", pureId="14992", uri="/dk/atira/pure/equipment/equipmenttypes/equipment") #TODO toca eliminar ese placeHolder de pureId
    category = ET.SubElement(equipment, "category", uri="/dk/atira/pure/equipment/category/classification") #TODO preguntarle a Francisco qué significa esto

    person_associations = ET.SubElement(equipment, "personAssociations")
    person_association = ET.SubElement(person_associations, "personAssociation")
    ET.SubElement(person_association, "personId").text = temp_eq_person_id

    org_units = ET.SubElement(equipment, "organisationalUnits")
    org_unit = ET.SubElement(org_units, "organisationalUnit", externalId="218")
    name = ET.SubElement(org_unit, "name", formatted="false")
    ET.SubElement(name, "text", locale="es_CO").text = "Facultad de Ingeniería" #TODO placeholder
    ET.SubElement(name, "text", locale="en_US").text = "Facultad de Ingeniería" #TODO placeholder



    # Descriptions
    descriptions = ET.SubElement(equipment, "descriptions")
    description = ET.SubElement(descriptions, "description", pureId=temp_eq_id) #TODO revisar qué poner en ese pureId
    value = ET.SubElement(description, "value", formatted="false")
    ET.SubElement(value, "text", locale="es_CO").text = temp_eq_es_description
    ET.SubElement(value, "text", locale="en_US").text = temp_eq_eng_description
    ET.SubElement(description, "type", uri="/dk/atira/pure/equipment/descriptions/equipmentdescription")

    # Equipment details
    equipment_details = ET.SubElement(equipment, "equipmentDetails")
    equipment_detail = ET.SubElement(equipment_details, "equipmentDetail", pureId=temp_eq_id) #TODO revisar qué poner en ese pureId
    name_detail = ET.SubElement(equipment_detail, "name", formatted="true")
    ET.SubElement(name_detail, "text", locale="es_CO").text = "Detalle en español"
    ET.SubElement(name_detail, "text", locale="en_US").text = "Detail in English"
    ET.SubElement(equipment_detail, "acquisitionDate").text = temp_eq_created
    # manufacturers = ET.SubElement(equipment_detail, "manufacturers")
    # manufacturer = ET.SubElement(manufacturers, "manufacturer", pureId=temp_eq_id) #TODO revisar qué poner en ese pureId
    # external_org_unit = ET.SubElement(manufacturer, "externalOrganisationalUnit", uuid="1df83989-a24d-440a-b847-df451e699af0") #TODO revisar qué poner en ese uuid
    # eou_name = ET.SubElement(external_org_unit, "name", formatted="false")
    # ET.SubElement(eou_name, "text", locale="es_CO").text = "Facultad de Ingeniería"
    
    # Managing Organizational Unit
    managing_org_unit = ET.SubElement(equipment, "managingOrganisationalUnit", externalId="218")
    name_org = ET.SubElement(managing_org_unit, "name", formatted="false")
    ET.SubElement(name_org, "text", locale="es_CO").text = "Facultad de Ingeniería"
    ET.SubElement(name_org, "text", locale="en_US").text = "Facultad de Ingeniería"

    # Contact Persons
    contact_persons = ET.SubElement(equipment, "contactPersons")
    contact_person = ET.SubElement(contact_persons, "contactPerson", externalId=temp_eq_person_id)
    name_contact = ET.SubElement(contact_person, "name", formatted="false")
    ET.SubElement(name_contact, "text").text = temp_eq_person

    # Addresses
    addresses = ET.SubElement(equipment, "addresses")
    address = ET.SubElement(addresses, "address", pureId=temp_eq_id)
    ET.SubElement(address, "addressType", uri="/dk/atira/pure/equipment/equipmentaddresstype/postal")
    ET.SubElement(address, "street").text = temp_eq_address
    ET.SubElement(address, "building").text = "Edificio José Gabriel Maldonado S.J.-Laboratorios"  # Placeholder
    ET.SubElement(address, "city").text = "Bogotá D.C."  # Placeholder
    ET.SubElement(address, "country", uri="/dk/atira/pure/core/countries/co")

    # Phone Numbers
    phone_numbers = ET.SubElement(equipment, "phoneNumbers")
    phone_number = ET.SubElement(phone_numbers, "phoneNumber", pureId=temp_eq_id) #TODO revisar qué poner en ese pureId
    ET.SubElement(phone_number, "value", formatted="true").text = str(temp_eq_phone)
    ET.SubElement(phone_number, "type", uri="/dk/atira/pure/equipment/equipmentphonenumbertype/phone")

    # Web Addresses
    if temp_eq_url != "":
        web_addresses = ET.SubElement(equipment, "webAddresses")
        web_address = ET.SubElement(web_addresses, "webAddress", pureId=temp_eq_id) #TODO revisar qué poner en ese pureId
        web_address_value = ET.SubElement(web_address, "value", formatted="false")
        ET.SubElement(web_address_value, "text", locale="es_CO").text = temp_eq_url
        ET.SubElement(web_address_value, "text", locale="en_US").text = temp_eq_url
        ET.SubElement(web_address, "type", uri="/dk/atira/pure/equipment/equipmentwebaddresstype/website")

    # Loan Type
    #set to lowercase and trim any space
    
    temp_eq_loan_type = temp_eq_loan_type.lower().strip()
    if temp_eq_loan_type == "ambos":
        temp_eq_loan_type = "/dk/atira/pure/equipment/loantypes/internalexternal"
    elif temp_eq_loan_type == "interno":
        temp_eq_loan_type = "/dk/atira/pure/equipment/loantypes/internal"
    else:
        temp_eq_loan_type = "/dk/atira/pure/equipment/loantypes/notavailable"

    loan_type = ET.SubElement(equipment, "loanType", uri=temp_eq_loan_type)
    
    # Keywords 
    keyword_groups = ET.SubElement(equipment, "keywordGroups")
    keyword_group = ET.SubElement(keyword_groups, "keywordGroup", logicalName="keywordContainers", pureId=temp_eq_id) #TODO revisar qué poner en ese pureId
    keyword_containers = ET.SubElement(keyword_group, "keywordContainers")
    keyword_container = ET.SubElement(keyword_containers, "keywordContainer", pureId=temp_eq_id) #TODO revisar qué poner en ese pureId
    free_keywords = ET.SubElement(keyword_container, "freeKeywords")
    
    # Split keywords and add them to the XML
    for keyword in temp_eq_keywords.split(','):
        free_keyword = ET.SubElement(free_keywords, "freeKeyword", locale="es_CO")
        free_keyword.text = keyword.strip()

# Convert the ElementTree to a string
tree = ET.ElementTree(result)
xml_str = ET.tostring(result, encoding="unicode", method="xml")

# Save to a file
with open(r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Resultado\2024_09_04_EQUIPOS_KV_V1.xml', "w", encoding="utf-8") as f:
    # Escribir la declaración manualmente
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    # Escribir el contenido del árbol XML
    f.write(xml_str)

print("xml generado")




