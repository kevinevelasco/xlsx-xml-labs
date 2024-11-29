"""
Autor: Kevin Estiven Velasco Pinto.
Este script procesa un archivo de Excel para generar un archivo XML estructurado. 
El propósito es convertir datos relacionados con equipos de infraestructura en un formato que pueda ser 
consumido por otros sistemas, específicamente Pure, siguiendo un esquema definido.

Módulos:
    - pandas: Para leer y manipular los datos en el archivo Excel.
    - spacy.language: Usado para integrarse con el detector de lenguaje (aunque no se utiliza aquí).
    - spacy_langdetect: Usado para detección de idiomas (aunque no se utiliza aquí).
    - xml.etree.ElementTree: Para construir y guardar la estructura XML.
    - datetime: Para manejar y formatear fechas.
    - re: Para realizar operaciones de expresiones regulares.

El proceso principal del código es:
    1. Leer un archivo Excel y convertirlo a un DataFrame de pandas.
    2. Limpiar y procesar los datos, asegurando que los valores ausentes sean reemplazados con valores por defecto.
    3. Iterar sobre cada equipo identificado por su `id_unico` para generar su respectiva estructura XML.
    4. Crear nodos XML con atributos y datos obtenidos del DataFrame, incluyendo validaciones específicas como la comprobación 
    de URLs y la conversión de fechas.
    5. Incorporar sub-elementos XML que representen detalles como nombres, descripciones, categorías, direcciones, teléfonos,
    personas asociadas, y más.

Variables Globales:
    - `result`: Elemento raíz del XML.
    - `items`: Subnodo de `result` que contendrá la información procesada.

Funciones:
    No hay funciones explícitas; el procesamiento ocurre directamente en el flujo principal del script.

Atributos específicos procesados del archivo Excel incluyen:
- Identificación única del equipo (`id_ubicacion`).
- Identificador y nombre del laboratorio (`id-lab`, `Espacio ubicación`).
- Información detallada del equipo:
  - Nombres y descripciones en español e inglés (`Nombre equipo español`, `Descripción español`, etc.).
  - Identificador de plaqueta de activos fijos.
  - Fechas relacionadas con adquisición y baja (`Fecha de adquisición`, `Fecha estimada baja`).
  - Responsable del equipo (`Id SIU responsable`, `nombre SIU responsable`).
  - Modelo, fabricante, país de fabricación.
  - Disponibilidad para préstamo, palabras clave, servicios y tipo de equipo.
- Relación con tipos de equipo y datos organizacionales.

El XML generado incluye elementos y sub-elementos organizados jerárquicamente para representar atributos como:
- `equipment`: Nodo raíz para cada equipo.
- `title`: Títulos en español e inglés del equipo.
- `type`: Tipo de equipo (e.g., equipo, instrumento, dispositivo).
- `descriptions`: Descripciones en diferentes idiomas.
- `equipmentDetails`: Detalles del equipo, incluyendo modelo y fabricante.
- `loanType`: Disponibilidad para préstamo interno, externo o ambos.
- `keywords`: Palabras clave relacionadas con el equipo.
- `services`: Servicios asociados al equipo.

Validaciones y transformaciones:
- Fechas de adquisición y baja se convierten al formato `YYYY-MM-DD` si son válidas.
- Descripciones y títulos se complementan con información adicional, como modelo y fabricante.
- Campos vacíos o inválidos, como palabras clave o descripciones no proporcionadas, son manejados adecuadamente.

Notas:
- Este código asume que el archivo Excel contiene todas las columnas necesarias. Si falta alguna columna, 
  podría generar errores.
- Algunos valores, como el tipo de equipo y su relación jerárquica, son placeholders que pueden requerir ajuste 
  según los datos reales.
- La estructura final del XML se guarda en la variable `result` y puede ser exportada o utilizada según sea necesario.
"""

import pandas as pd
from spacy.language import Language
from spacy_langdetect import LanguageDetector
import xml.etree.ElementTree as ET
import datetime
import re

# Crear el nodo raíz del XML
"""
Inicializa la estructura XML con los nodos `result` e `items`. Se configuran atributos de esquema.
"""
result = ET.Element("result")
result.set('xmlns:xsi',"http://www.w3.org/2001/XMLSchema-instance")
result.set('xsi:schemaLocation',"http://localhost:8080/ws/api/524/xsd/schema1.xsd")
items = ET.SubElement(result,"items")

# Leer el archivo de Excel y convertirlo en un DataFrame
"""
Carga datos desde un archivo de Excel y realiza operaciones preliminares de limpieza, como rellenar valores nulos.
"""
dataframe1 = pd.read_excel(
    r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Excel\Formato Infraestructura - Perfiles y Capacidades - Ajustado.xlsx', sheet_name='Equipos', dtype=object)
dataframe1 = dataframe1.fillna('')
column = dataframe1.id_ubicacion.unique()

# Procesar cada ubicación única
"""
Itera sobre las ubicaciones únicas en el DataFrame y genera nodos XML basados en los datos relacionados.
"""
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

    # Validación y formateo de descripciones
    """
    Maneja las descripciones en español e inglés, verificando si son válidas y ajustándolas según el modelo.
    """
    if (isinstance(temp_eq_es_description,str)):
        temp_eq_es_description = temp_eq_es_description
    else:
        temp_eq_es_description = ""
    
    if (isinstance(temp_eq_eng_description,str)):
        temp_eq_eng_description = temp_eq_eng_description
    else:
        temp_eq_eng_description = ""

    # Formatear fechas
    """
    Convierte fechas de adquisición y baja a formato `YYYY-MM-DD`, si son válidas.
    """
    if isinstance(temp_eq_adquisicion, datetime.datetime):
        temp_eq_adquisicion = temp_eq_adquisicion.strftime("%Y-%m-%d")
    else:
        temp_eq_adquisicion = ""

    if isinstance(temp_eq_baja, datetime.datetime):
        temp_eq_baja = temp_eq_baja.strftime("%Y-%m-%d")
    else: 
        temp_eq_baja = ""

    # Validación de otros campos
    """
    Limpia y verifica el contenido de los campos del modelo y fabricante.
    """
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

    # Crear nodos XML para el equipo
    """
    Construye el nodo XML `equipment` y sus subnodos, incorporando títulos, descripciones, detalles y asociaciones.
    """
    equipment= ET.SubElement(items,"equipment", externalId=temp_eq_id)

    # Añadir títulos en múltiples idiomas
    title = ET.SubElement(equipment, "title", formatted="true")
    ET.SubElement(title, "text", locale="es_CO").text = temp_eq_es_name
    ET.SubElement(title, "text", locale="en_US").text = temp_eq_eng_name

    # Añadir tipo de equipo
    type = ET.SubElement(equipment, "type", uri="/dk/atira/pure/equipment/equipmenttypes/equipment/" + temp_eq_type)
    type_term = ET.SubElement(type, "term", formatted = "false")
    ET.SubElement(type_term, "text", locale="en_US").text = "Equipment"
    ET.SubElement(type_term, "text", locale="es_CO").text = "Equipo"

    # Añadir categoría
    category = ET.SubElement(equipment, "category", uri="/dk/atira/pure/equipment/category") 

    # Añadir descripciones
    descriptions = ET.SubElement(equipment, "descriptions")
    description = ET.SubElement(descriptions, "description")
    value = ET.SubElement(description, "value", formatted="false")
    ET.SubElement(value, "text", locale="es_CO").text = temp_eq_es_description
    ET.SubElement(value, "text", locale="en_US").text = temp_eq_eng_description
    description_type = ET.SubElement(description, "type", uri="/dk/atira/pure/equipment/descriptions/equipmentdescription")
    description_type_term = ET.SubElement(description_type, "term", formatted="false")
    ET.SubElement(description_type_term, "text", locale="es_CO").text = "Descripción del equipo"
    ET.SubElement(description_type_term, "text", locale="en_US").text = "Equipment description"


    # Añadir detalles del equipo
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

        #Validar si la fecha de adquisición, baja no están vacías, para asignarlas 
        if temp_eq_adquisicion != "":
            ET.SubElement(equipment_detail, "acquisitionDate").text = temp_eq_adquisicion
        if temp_eq_baja != "":
            ET.SubElement(equipment_detail, "decommissionDate").text = temp_eq_baja
        if temp_eq_fabricante != "":

            #Si no está vacío el fabricante, se procede a llenar sus detalles
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

    # Información sobre las personas y sus asociaciones
    person_associations = ET.SubElement(equipment, "personAssociations")
    person_association = ET.SubElement(person_associations, "personAssociation")
    ET.SubElement(person_association, "person", externalId=str(temp_eq_siu))
    ET.SubElement(person_association, "personRole",uri="/dk/atira/pure/equipment/roles/equipment/manager")
    person_org_units = ET.SubElement(person_association, "organisationalUnits")
    person_org_unit = ET.SubElement(person_org_units, "organisationalUnit", externalId="218")
    
    # Unidades organizacionales de los equipos
    org_units = ET.SubElement(equipment, "organisationalUnits")
    org_unit = ET.SubElement(org_units, "organisationalUnit", externalId="218")
    name = ET.SubElement(org_unit, "name", formatted="false")
    ET.SubElement(name, "text", locale="es_CO").text = "Facultad de Ingeniería"
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

    # Personas de contacto
    contact_persons = ET.SubElement(equipment, "contactPersons")
    contact_person = ET.SubElement(
        contact_persons, "contactPerson", externalId=str(temp_eq_siu), externalIdSource="synchronisedPerson",externallyManaged="true")
    name_contact = ET.SubElement(contact_person, "name", formatted="false")
    ET.SubElement(name_contact, "text").text = temp_eq_nombre_siu


    # Tipo de préstamo
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

    #Parents de los Equipos, es decir los Laboratorios
    parents_group = ET.SubElement(equipment, "parents")
    parent = ET.SubElement(parents_group, "parent", externalId=temp_eq_id_lab)
    parent_name = ET.SubElement(parent, "name", formatted="false")
    ET.SubElement(parent_name, "text", locale="es_CO").text = temp_eq_lab_es_name
    parent_type = ET.SubElement(parent, "type", uri="/dk/atira/pure/equipment/equipmenttypes/equipment/" + temp_eq_parent_type)
    parent_type_term = ET.SubElement(parent_type, "term", formatted="false")
    ET.SubElement(parent_type_term, "text", locale="es_CO").text = "Laboratorio"
    ET.SubElement(parent_type_term, "text", locale="en_US").text = "Laboratory"
    
    # Palabras clave en español e inglés
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
    
    # Por cada palabra clave del Excel, se añade en el XML
    for keyword in temp_eq_keywords.split(','):
        free_keyword = ET.SubElement(free_keywords_list, "freeKeyword")
        free_keyword.text = keyword.strip()    

# Crear el árbol XML y guardar en un archivo
tree = ET.ElementTree(result)
xml_str = ET.tostring(result, encoding="unicode", method="xml")

# Guardar el archivo XML
with open(r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Resultado\2024_11_07_EQUIPOS_KV_V6.xml', "w", encoding="utf-8") as f:
    # Escribir la declaración manualmente
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    # Escribir el contenido del árbol XML
    f.write(xml_str)

print("XML generado")

# Fin del script
# Nota: Este código es un ejemplo simplificado y puede requerir ajustes adicionales para adaptarse a necesidades específicas.
'''
Al finalizar el script se genera un archivo XML con la estructura y los datos procesados.
Todos los archivos generados se encuentran en la carpeta .../Infraestructura/Resultado/.
'''




