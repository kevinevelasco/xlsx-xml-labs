# importing required classes
import os
from pypdf import PdfReader
import re
import pandas as pd

# Función para corregir los espacios antes de las letras acentuadas
def corregir_acentuacion(texto):
    # Expresión regular para buscar un espacio antes de una vocal acentuada (mayúscula)
    texto_corregido = re.sub(r'(\w)\s([ÁÉÍÓÚ])', r'\1\2', texto)
    return texto_corregido

pdf_folder = r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Registros de Software\PDFs'

# Lista para acumular todos los datos procesados
all_data = []

# Obtener todos los archivos PDF de la carpeta
pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]

# Procesar cada archivo PDF
for pdf_file in pdf_files:
    pdf_path = os.path.join(pdf_folder, pdf_file)
    
    # Crear un objeto PdfReader para leer el archivo actual
    reader = PdfReader(pdf_path)

    text = ""
    for i in range(len(reader.pages)):
        page_text = reader.pages[i].extract_text()
        text += page_text

    text = corregir_acentuacion(text)

    #print(text)
    # Inicializar variables para los datos
    titulo_original = ""
    anio_creacion = None
    anio_edicion = None
    pais_origen = ""
    detalles_obra = {}
    descripcion_obra = ""
    observaciones_obra = ""

    # Bandera para identificar que estamos en las secciones correctas
    in_datos_obra = False
    in_descripcion_obra = False
    in_observaciones_obra = False

    lines = text.split('\n')

    for i, line in enumerate(lines):
        line = line.strip()  # Limpiar espacios en blanco

        # Detectar el inicio de "2. DATOS DE LA OBRA"
        if "2. DATOS DE LA OBRA" in line:
            in_datos_obra = True
            continue

        # Si estamos en la sección de datos de la obra
        if in_datos_obra:
            # Extraer "Título Original"
            if "Título Original" in line:
                titulo_original = line.split("Título Original")[-1].strip() + " "  # Guardar la primera línea del título
                # Capturar todas las líneas siguientes hasta encontrar "Año de Creación"
                for j in range(i + 1, len(lines)):
                    next_line = lines[j].strip()
                    if "Año de Creación" in next_line:
                        break
                    titulo_original += next_line + " "  # Continuar añadiendo líneas al título

                titulo_original = titulo_original.strip()  # Limpiar cualquier espacio extra
            # Extraer "Año de Creación", "Año Edición", y "País de Origen"
            if "Año de Creación" in line and "País de Origen" in line:
                # Encontrar los años
                parts = re.findall(r'\d{4}', line)
                if len(parts) > 0:
                    anio_creacion = parts[0]  # Primer año
                if len(parts) > 1:
                    anio_edicion = parts[1]  # Segundo año si existe

                # Extraer el país de origen (entre "Edición" y "País de Origen")
                pais_origen_match = re.search(r'Edición\s+(\w+)\s+País de Origen', line)
                if pais_origen_match:
                    pais_origen = pais_origen_match.group(1)

            # Extraer "CLASE DE OBRA"
            if "CLASE DE OBRA" in line:
                # Añadir o actualizar si ya existe la clase de obra
                if "CLASE DE OBRA" in detalles_obra:
                    detalles_obra["CLASE DE OBRA"] += f", {line.split('CLASE DE OBRA')[-1].strip()}"
                else:
                    detalles_obra["CLASE DE OBRA"] = line.split("CLASE DE OBRA")[-1].strip()

            # Extraer "CARACTER DE LA OBRA"
            if "CARACTER DE LA OBRA" in line:
                # Añadir o actualizar si ya existe el carácter de la obra
                if "CARACTER DE LA OBRA" in detalles_obra:
                    detalles_obra["CARACTER DE LA OBRA"] += f", {line.split('CARACTER DE LA OBRA')[-1].strip()}"
                else:
                    detalles_obra["CARACTER DE LA OBRA"] = line.split("CARACTER DE LA OBRA")[-1].strip()

            # Extraer "ELEMENTOS APORTADOS DE SOPORTE LOGICO"
            if "ELEMENTOS APORTADOS DE SOPORTE LOGICO" in line:
                # Añadir o actualizar si ya existen elementos aportados
                if "ELEMENTOS APORTADOS DE SOPORTE LOGICO" in detalles_obra:
                    detalles_obra["ELEMENTOS APORTADOS DE SOPORTE LOGICO"] += f", {line.split('ELEMENTOS APORTADOS DE SOPORTE LOGICO')[-1].strip()}"
                else:
                    detalles_obra["ELEMENTOS APORTADOS DE SOPORTE LOGICO"] = line.split("ELEMENTOS APORTADOS DE SOPORTE LOGICO")[-1].strip()

        # Detectar el inicio de "3. DESCRIPCIÓN DE LA OBRA"
        if "3. DESCRIPCIÓN DE LA OBRA" in line or "3. DESCRICIPCIÓN DE LA OBRA" in line:
            in_descripcion_obra = True
            continue

        # Capturar la descripción de la obra
        if in_descripcion_obra:
            # Terminar la captura si encontramos la siguiente sección
            if "4. OBSERVACIONES GENERALES DE LA OBRA" in line:
                in_descripcion_obra = False
                continue  # Salir para no capturar esta línea

            # Añadir todas las líneas a la descripción de la obra solo si estamos en la sección correcta
            descripcion_obra += line + " "


        # Capturar las observaciones generales de la obra
        if in_observaciones_obra:
            if "5. DATOS DEL SOLICITANTE" in line:
                break  # Finalizar la captura cuando llegamos a "5. DATOS DEL SOLICITANTE"
            observaciones_obra += line + " "

    # Limpiar espacios adicionales
    titulo_original = titulo_original.strip()
    descripcion_obra = descripcion_obra.strip()
    observaciones_obra = observaciones_obra.strip()

    # Crear un diccionario con los datos extraídos
    obra_data = {
        "Título Original": titulo_original,
        "Año de Creación": anio_creacion,
        "Año de Edición": anio_edicion,
        "País de Origen": pais_origen,
        "Detalles de la Obra": detalles_obra,
        "Descripción de la Obra": descripcion_obra,
        "Observaciones Generales de la Obra": observaciones_obra
    }

    # Imprimir los datos extraídos
    print(obra_data)

    # Inicializar variables
    libro = tomo = partida = fecha_registro = None

    # Buscar la línea con "Fecha Registro" y extraer los valores
    for i, line in enumerate(lines):
        if "Fecha Registro" in line:
            # Extraer los números de Libro, Tomo y Partida de la línea
            # Buscamos la secuencia de tres números separados por guiones
            libro_tomo_partida = line.split("Fecha Registro")[-1].strip()  # Tomar lo que está después de "Fecha Registro"
            libro, tomo, partida = libro_tomo_partida.split('-')
            
            # La fecha de registro debería estar en la siguiente línea
            fecha_registro = lines[i + 1].strip()
            break

    # Crear el diccionario con los datos extraídos
    registro_data = {
        "Libro": libro,
        "Tomo": tomo,
        "Partida": partida,
        "Fecha Registro": fecha_registro
    }

    print(registro_data)

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
            data["Nombres y Apellidos"] = razon_social

            # Obtener el Nit y eliminar "TITULAR" o "PRODUCTOR"
            nit_info = re.findall(r"Nit\s+(\d+[A-Z]*)", entry)
            if nit_info:
                nit = re.sub(r'(TITULAR|PRODUCTOR)', '', nit_info[0]).strip()
                data["Número de Identificación"] = "Nit" + nit

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

    # Simulación de datos procesados (reemplaza con los datos extraídos reales)
    for person in processed_data:
        data = {
            "libro": libro,
            "tomo": tomo,
            "partida": partida,
            "Fecha Registro": fecha_registro,
            "AUTOR_tipo": person.get("Tipo", ""),
            "AUTOR_nombre": person.get("Nombres y Apellidos", ""),
            "AUTOR_identificacion": person.get("Número de Identificación", ""),
            "AUTOR_pais": person.get("País", ""),
            "AUTOR_direccion": person.get("Dirección", ""),
            "AUTOR_ciudad": person.get("Ciudad", ""),
            "OBRA_titulo_original": titulo_original,
            "OBRA_año_creación": anio_creacion,
            "OBRA_país_origen": pais_origen,
            "OBRA_año_edicion": anio_edicion,
            "OBRA_clase_de_obra": ", ".join(detalles_obra.get("CLASE DE OBRA", "").split(", ")),  # Separar por comas
            "OBRA_caracter_de_la_obra": ", ".join(detalles_obra.get("CARACTER DE LA OBRA", "").split(", ")),  # Separar por comas
            "OBRA_elementos_aportados_de_soporte_lógico": ", ".join(detalles_obra.get("ELEMENTOS APORTADOS DE SOPORTE LOGICO", "").split(", ")),  # Separar por comas
            "OBRA_descripcion": descripcion_obra,
            "OBRA_observaciones": observaciones_obra,
        }

        # Añadir los datos de esta persona a la lista
        all_data.append(data)

# Convertir todos los datos acumulados en un DataFrame
df = pd.DataFrame(all_data)

# Especificar la ruta de guardado del archivo Excel
output_path = os.path.join(r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Registros de Software\Excel', 'data.xlsx')

# Guardar el DataFrame en un archivo Excel
df.to_excel(output_path, index=False)

print(f"Datos de {len(pdf_files)} archivos PDF guardados en: {output_path}")