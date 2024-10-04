# importing required classes
from pypdf import PdfReader
import re

# creating a pdf reader object
reader = PdfReader(r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Registros de Software\PDFs\APP CODIFICO 2.0.pdf')

text = ""
for i in range(len(reader.pages)):
    page_text = reader.pages[i].extract_text()
    text += page_text

# Extracting text from "1. DATOS DE LAS PERSONAS" to "2. DATOS DE LA OBRA"
start = text.find("1. DATOS DE LAS PERSONAS")
end = text.find("2. DATOS DE LA OBRA")
if start != -1:
    start += len("1. DATOS DE LAS PERSONAS")

extracted_text = text[start:end]

# If the text contains "Fecha Registro*", delete that line and the next three lines
lines = extracted_text.split("\n")
for i in range(len(lines)):
    if "Fecha Registro" in lines[i]:
        del lines[i:i+4]
        break

# Join the remaining lines
extracted_text = "\n".join(lines)

# Split text by lines
lines = extracted_text.split("\n")

# Initialize an empty list to store the data
persons_data = []
current_entry = []

# Iterate over lines and store the blocks of text associated with each "Dirección"
for line in lines:
    line = line.strip()  # Clean the line
    if line.startswith("Dirección"):
        # Add the line with the current address to the ongoing entry
        current_entry.append(line)
        # Save the completed entry and reset for the next one
        persons_data.append("\n".join(current_entry))
        current_entry = []
    else:
        current_entry.append(line)  # Continue adding lines until we find a new "Dirección"

# Function to process each entry and extract relevant data
def process_entry(entry):
    data = {}
    lines = entry.split("\n")
    
    # Limpiar las líneas en blanco o irrelevantes
    lines = [line.strip() for line in lines if line.strip()]
    
    # Unir líneas adyacentes si el número de identificación o Nit se encuentra en la siguiente línea
    full_entry = " ".join(lines)

    # Buscar "No de identificación" o "Razón Social" y procesar
    if "No de identificación" in full_entry:
        id_info = full_entry.split("No de identificación")
        
        # Eliminar "Nombres y Apellidos" del valor
        nombres_y_apellidos = id_info[0].replace("Nombres y Apellidos", "").strip()
        data["Nombres y Apellidos"] = nombres_y_apellidos
        
        id_number_and_type = id_info[1].split()

        # Concatenar tipo y número de identificación, y eliminar "AUTOR*" si está presente
        if len(id_number_and_type) >= 2:
            id_number = f"{id_number_and_type[0]}{id_number_and_type[1]}"
        else:
            id_number = id_number_and_type[0]
        
        # Eliminar todo lo que esté después de "AUTOR" en el número de identificación
        if "AUTOR" in id_number:
            id_number = id_number.split("AUTOR")[0].strip()
        
        data["Número de Identificación"] = id_number

        # Extraer el tipo (AUTOR, PRODUCTOR, etc.)
        for tipo in ["AUTOR", "PRODUCTOR", "TITULAR DERECHO PATRIMONIAL"]:
            if tipo in full_entry:
                data["Tipo"] = tipo
                break

        # Razón Social para las empresas
    if "Razón Social" in entry:
        # Capturar la razón social que está antes de "Razón Social"
        razon_social = entry.split("Razón Social")[0].strip()
        data["Razón Social"] = razon_social

        # Obtener el Nit y eliminar "TITULAR" o "PRODUCTOR"
        nit_info = re.findall(r"Nit\s+(\d+[A-Z]*)", entry)
        if nit_info:
            nit = re.sub(r'(TITULAR|PRODUCTOR)', '', nit_info[0]).strip()
            data["Nit"] = nit

        # Extraer el tipo de empresa (TITULAR, PRODUCTOR, etc.)
        tipo_info = re.search(r'(TITULAR DERECHO PATRIMONIAL|PRODUCTOR)', entry)
        if tipo_info:
            data["Tipo"] = tipo_info.group()

    # Extraer el país (solo la primera palabra)
    if "Nacional de" in full_entry:
        data["País"] = full_entry.split("Nacional de")[1].strip().split()[0]

    # Extraer la dirección y eliminar todo lo que esté después de "Ciudad"
    if "Dirección" in full_entry:
        direccion = full_entry.split("Dirección")[1].strip()
        if "Ciudad" in direccion:
            data["Dirección"] = direccion.split("Ciudad")[0].strip()
        else:
            data["Dirección"] = direccion

    # Extraer la ciudad
    if "Ciudad:" in full_entry:
        data["Ciudad"] = full_entry.split("Ciudad:")[1].strip()

    return data

# Process each entry in persons_data
processed_data = []
for entry in persons_data:
    processed_data.append(process_entry(entry))

# Print the processed data
for data in processed_data:
    print(data)
    print("-----")

print("El tamaño del arreglo es:", len(persons_data))

