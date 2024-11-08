import pandas as pd
from spacy.language import Language
from spacy_langdetect import LanguageDetector
import xml.etree.ElementTree as ET
import datetime
import re

#Create Element Tree root class
result = ET.Element("result")
result.set('xmlns:xsi',"http://www.w3.org/2001/XMLSchema-instance")
result.set('xsi:schemaLocation',"http://localhost:8080/ws/api/524/xsd/schema1.xsd")
items = ET.SubElement(result,"items")

# Read excel and convert to dataframe
dataframe1 = pd.read_excel(
    r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Excel\Formato Infraestructura - Perfiles y Capacidades - Ajustado.xlsx', sheet_name='Equipos', dtype=object)
dataframe1 = dataframe1.fillna('')
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
    temp_eq_nombre_siu = temp_eq["nombre SIU responsable"].values[0]
    temp_eq_model = temp_eq["Modelo"].values[0]
    temp_eq_fabricante = temp_eq["Nombre fabricante"].values[0]
    temp_eq_pais_fabricante = temp_eq["Pais fabricante"].values[0]
    temp_eq_loan_type = temp_eq["Disponible para prestamo (interno/externo/ambos)"].values[0]
    temp_eq_keywords = temp_eq["Palabras clave"].values[0]
    temp_eq_services = temp_eq["Servicios"].values[0]
    temp_eq_type = temp_eq["tipo"].values[0]
    temp_eq_parent_type = temp_eq["tipo-parent"].values[0]

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
    else:
        temp_eq_adquisicion = ""

    if isinstance(temp_eq_baja, datetime.datetime):
        temp_eq_baja = temp_eq_baja.strftime("%Y-%m-%d")
    else: 
        temp_eq_baja = ""


    if isinstance(temp_eq_model, str):
        temp_eq_model = temp_eq_model
    else:
        temp_eq_model = ""

    if temp_eq_model == "No aporta" or temp_eq_model == "-" or temp_eq_model == "":
        temp_eq_model = ""
    else:
        temp_eq_model = temp_eq_model


    if isinstance(temp_eq_fabricante, str):
        temp_eq_model = temp_eq_model
    else:
        temp_eq_model = ""

    #si el nombre es No aporta, -, o vacío, se le asigna "" para que quede vacío
    if temp_eq_fabricante == "No aporta" or temp_eq_fabricante == "-" or temp_eq_fabricante == "":
        temp_eq_fabricante = ""
    else:
        temp_eq_fabricante = temp_eq_fabricante

    if isinstance(temp_eq_services, str):
        temp_eq_services = temp_eq_services
    else:
        temp_eq_services = ""

    if(temp_eq_model != ""):    
        temp_eq_es_description = "Modelo: " + str(temp_eq_model) + ". " + temp_eq_es_description
        temp_eq_eng_description = "Model: " + str(temp_eq_model) + ". " + temp_eq_eng_description
        temp_eq_es_name = temp_eq_es_name + ". Modelo: " + str(temp_eq_model)
        temp_eq_eng_name = temp_eq_eng_name + ". Model: " + str(temp_eq_model)

    equipment= ET.SubElement(items,"equipment", externalId=temp_eq_id)

    title = ET.SubElement(equipment, "title", formatted="true")
    ET.SubElement(title, "text", locale="es_CO").text = temp_eq_es_name
    ET.SubElement(title, "text", locale="en_US").text = temp_eq_eng_name

    type = ET.SubElement(equipment, "type", uri="/dk/atira/pure/equipment/equipmenttypes/equipment/" + temp_eq_type)
    type_term = ET.SubElement(type, "term", formatted = "false")
    ET.SubElement(type_term, "text", locale="en_US").text = "Equipment"
    ET.SubElement(type_term, "text", locale="es_CO").text = "Equipo"

    category = ET.SubElement(equipment, "category", uri="/dk/atira/pure/equipment/category") 

    descriptions = ET.SubElement(equipment, "descriptions")
    description = ET.SubElement(descriptions, "description")
    value = ET.SubElement(description, "value", formatted="false")
    ET.SubElement(value, "text", locale="es_CO").text = temp_eq_es_description
    ET.SubElement(value, "text", locale="en_US").text = temp_eq_eng_description
    description_type = ET.SubElement(description, "type", uri="/dk/atira/pure/equipment/descriptions/equipmentdescription")
    description_type_term = ET.SubElement(description_type, "term", formatted="false")
    ET.SubElement(description_type_term, "text", locale="es_CO").text = "Descripción del equipo"
    ET.SubElement(description_type_term, "text", locale="en_US").text = "Equipment description"


    # Equipment details
    if temp_eq_adquisicion and temp_eq_baja and temp_eq_fabricante != "":
        id = re.sub('[^A-Za-z0-9]+', '', str(temp_eq_plaqueta))
        equipment_details = ET.SubElement(equipment, "equipmentDetails")
        equipment_detail = ET.SubElement(equipment_details, "equipmentDetail", pureId=id)
        name_detail = ET.SubElement(equipment_detail, "name", formatted="true")
        ET.SubElement(name_detail, "text", locale="es_CO").text = temp_eq_es_name
        ET.SubElement(name_detail, "text", locale="en_US").text = temp_eq_eng_name

        ids_detail = ET.SubElement(equipment_detail, "ids")
        ids_detail_id = ET.SubElement(ids_detail, "id", pureId=id)
        ET.SubElement(ids_detail_id, "value", formatted="false").text = str(temp_eq_plaqueta)
        type_detail = ET.SubElement(ids_detail_id, "type", uri="/dk/atira/pure/equipment/equipmentsources/internalid")
        type_detail_term = ET.SubElement(type_detail, "term", formatted="false")
        ET.SubElement(type_detail_term, "text", locale="es_CO").text = "ID interna"
        ET.SubElement(type_detail_term, "text", locale="en_US").text = "Internal ID"

        if temp_eq_adquisicion != "":
            ET.SubElement(equipment_detail, "acquisitionDate").text = temp_eq_adquisicion
        if temp_eq_baja != "":
            ET.SubElement(equipment_detail, "decommissionDate").text = temp_eq_baja
        if temp_eq_fabricante != "":
            id = re.sub('[^A-Za-z0-9]+', '', str(temp_eq_plaqueta))
            manufacturers = ET.SubElement(equipment_detail, "manufacturers")
            manufacturer = ET.SubElement(manufacturers, "manufacturer", pureId=id)
            external_org_unit = ET.SubElement(manufacturer, "externalOrganisationalUnit")
            eou_name = ET.SubElement(external_org_unit, "name", formatted="false")
            ET.SubElement(eou_name, "text", locale="es_CO").text = temp_eq_fabricante
            eou_type = ET.SubElement(external_org_unit, "type", uri="/dk/atira/pure/ueoexternalorganisation/ueoexternalorganisationtypes/ueoexternalorganisation/unknown") #TODO revisar si la URI está bien
            eou_type_term = ET.SubElement(eou_type, "term", formatted="false")
            ET.SubElement(eou_type_term, "text", locale="es_CO").text = "Sin especificar"
            ET.SubElement(eou_type_term, "text", locale="en_US").text = "Unknown"
            #TODO preguntar a Francisco cómo poner el país del fabricante o si omitirlo

    person_associations = ET.SubElement(equipment, "personAssociations")
    person_association = ET.SubElement(person_associations, "personAssociation")
    ET.SubElement(person_association, "person", externalId=str(temp_eq_siu))
    ET.SubElement(person_association, "personRole",uri="/dk/atira/pure/equipment/roles/equipment/manager")
    person_org_units = ET.SubElement(person_association, "organisationalUnits")
    person_org_unit = ET.SubElement(person_org_units, "organisationalUnit", externalId="218")
    

    org_units = ET.SubElement(equipment, "organisationalUnits")
    org_unit = ET.SubElement(org_units, "organisationalUnit", externalId="218")
    name = ET.SubElement(org_unit, "name", formatted="false")
    # TODO placeholder
    ET.SubElement(name, "text", locale="es_CO").text = "Facultad de Ingeniería"
    # TODO placeholder
    ET.SubElement(name, "text", locale="en_US").text = "School of Engineering"
    
    # Managing Organizational Unit
    #TODO todo aquí es un placeholder de la Facultad de Ingeniería
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
        contact_persons, "contactPerson", externalId=str(temp_eq_siu), externalIdSource="synchronisedPerson",externallyManaged="true")
    name_contact = ET.SubElement(contact_person, "name", formatted="false")
    ET.SubElement(name_contact, "text").text = temp_eq_nombre_siu


    # Loan Type
    #set to lowercase and trim any space
    
    temp_eq_loan_type = temp_eq_loan_type.lower().strip()
    loan_type_spanish = ""
    loan_type_english = ""
    if temp_eq_loan_type == "ambos":
        temp_eq_loan_type = "/dk/atira/pure/equipment/loantypes/internalexternal"
        loan_type_spanish = "Disponible para el préstamo, a nivel interno y externo"
        loan_type_english = "Available for loan - internal and external"
    elif temp_eq_loan_type == "interno":
        temp_eq_loan_type = "/dk/atira/pure/equipment/loantypes/internal"
        loan_type_spanish = "Disponible para el préstamo, solo a nivel interno"
        loan_type_english = "Available for loan - internal only"
    else:
        temp_eq_loan_type = "/dk/atira/pure/equipment/loantypes/notavailable"
        loan_type_spanish = "No está disponible para el préstamo"
        loan_type_english = "Not available for loan"

    loan_type = ET.SubElement(equipment, "loanType", uri=temp_eq_loan_type)
    loan_type_term = ET.SubElement(loan_type, "term", formatted="false")
    ET.SubElement(loan_type_term, "text", locale="en_US").text = loan_type_english
    ET.SubElement(loan_type_term, "text", locale="es_CO").text = loan_type_spanish

    #si los servicios contienen | lo convierte en coma
    temp_lab_services = temp_eq_services.replace("|",",")

    if temp_lab_services != "":
        loan_term = ET.SubElement(equipment, "loanTerm", formatted="true")
        loan_term_text = ET.SubElement(loan_term, "text", locale="es_CO")
        loan_term_text.text = temp_lab_services

    #parents
    parents_group = ET.SubElement(equipment, "parents")
    parent = ET.SubElement(parents_group, "parent", externalId=temp_eq_id_lab)
    parent_name = ET.SubElement(parent, "name", formatted="false")
    ET.SubElement(parent_name, "text", locale="es_CO").text = temp_eq_lab_es_name

    parent_type = ET.SubElement(parent, "type", uri="/dk/atira/pure/equipment/equipmenttypes/equipment/" + temp_eq_parent_type)
    parent_type_term = ET.SubElement(parent_type, "term", formatted="false")
    ET.SubElement(parent_type_term, "text", locale="es_CO").text = "Laboratorio"
    ET.SubElement(parent_type_term, "text", locale="en_US").text = "Laboratory"
    
    # Keywords
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
    for keyword in temp_eq_keywords.split(','):
        free_keyword = ET.SubElement(free_keywords_list, "freeKeyword")
        free_keyword.text = keyword.strip()    

# Convert the ElementTree to a string
tree = ET.ElementTree(result)
xml_str = ET.tostring(result, encoding="unicode", method="xml")

# Save to a file
with open(r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Resultado\2024_11_07_EQUIPOS_KV_V6.xml', "w", encoding="utf-8") as f:
    # Escribir la declaración manualmente
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    # Escribir el contenido del árbol XML
    f.write(xml_str)

print("xml generado")




