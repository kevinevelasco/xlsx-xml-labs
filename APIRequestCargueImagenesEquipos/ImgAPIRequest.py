"""
Autor: Kevin Estiven Velasco Pinto.
Este script procesa datos de equipos desde un archivo Excel, asocia imágenes con los equipos utilizando un identificador específico, 
y carga los datos, incluidas las imágenes codificadas, a una API externa mediante solicitudes HTTP.

Módulos utilizados:
-------------------
- `base64`: Para codificar imágenes en formato Base64 para su carga.
- `requests`: Para realizar solicitudes HTTP a la API externa.
- `pandas`: Para manejar datos de archivos Excel y gestionar estructuras de datos.
- `os`: Para manipulaciones de rutas de archivos y verificación de existencia de archivos.
- `json`: Para manejar datos en formato JSON.

Dependencias:
--------------
Asegúrese de que los siguientes módulos estén instalados:
- `requests`
- `pandas`
- `openpyxl` (para leer archivos Excel)

Uso:
-----
Reemplace las rutas de los archivos y la clave API en `base_url`, `headers` y las rutas de archivo según su configuración.
"""
import base64
from requests.auth import HTTPBasicAuth
import requests
import pandas as pd
import os
import json

# URL base del endpoint de la API de Pure
base_url = 'https://puj-staging.elsevierpure.com/ws/api/equipment'

# Headers para los HTTP requests
headers = {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'api-key': ''
}

# Leer el Excel con la información a enviar a la API
dataframe1 = pd.read_excel(
    r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Excel\Formato Infraestructura - Perfiles y Capacidades - Ajustado.xlsx',
    sheet_name='uuids', 
    dtype=object
)

# Extraer UUIDs únicos del archivo Excel
column = dataframe1.UUID.unique()

# Ruta de la carpeta que contiene las imágenes de los equipos
image_folder_path = r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Fotos'

# Formatos de imagen válidos
image_formats = ['.png', '.jpg', '.jpeg', '.gif']

# Lista para almacenar los datos de los equipos
equipos = []

# Iterar sobre los UUIDs únicos y extraer información de los equipos
for uid in column:
    temp_eq = dataframe1[dataframe1["UUID"] == uid]
    temp_eq_id = str(uid)
    temp_eq_source_id = temp_eq["Source ID"].values[0]
    temp_eq_plaqueta = temp_eq["Plaqueta"].values[0]
    temp_eq_nombre = temp_eq["Nombre Equipo"].values[0]

    # Buscar la imagen correspondiente al equipo
    temp_eq_image_path = None
    for fmt in image_formats:
        image_path = os.path.join(image_folder_path, f"{temp_eq_plaqueta}{fmt}")
        if os.path.exists(image_path):
            temp_eq_image_path = image_path
            break

    # Agregar los datos del equipo a la lista
    equipos.append({
        "Nombre Equipo": temp_eq_nombre,
        "UUID": temp_eq_id,
        "Source ID": temp_eq_source_id,
        "Plaqueta": temp_eq_plaqueta,
        "Image Path": temp_eq_image_path  # Path to the image or None if not found
    })

# Crear un DataFrame con los datos de los equipos
equipos_df = pd.DataFrame(equipos)

# Filtrar los equipos con imágenes válidas
imagenes_con_ruta = equipos_df[equipos_df['Image Path'].notna()]

# Mostrar la cantidad de imágenes con rutas y los detalles de los equipos
cantidad_imagenes_con_ruta = imagenes_con_ruta.shape[0]
print(f"Cantidad de imágenes con ruta: {cantidad_imagenes_con_ruta}")
print("Imágenes con ruta:")
print(imagenes_con_ruta[['UUID', 'Source ID', 'Plaqueta', 'Image Path']])

# Iterar sobre los equipos con imágenes y cargar las imágenes a la API
for index, row in imagenes_con_ruta.iterrows():
    first_image_path = row['Image Path']
    first_image_filename = os.path.basename(first_image_path)
    mime_type = "image/" + first_image_filename.split('.')[-1]  # Derive MIME type

    # Leer y codificar la imagen en formato Base64
    with open(first_image_path, "rb") as image_file:
        encoded_path = base64.b64encode(image_file.read()).decode('utf-8')

    # Construir el payload JSON para la solicitud PUT
    data = {
        "images": [{
            "fileName": first_image_filename,
            "mimeType": mime_type,
            "size": os.path.getsize(first_image_path),
            "uploadedFile": {
                "digest": mime_type,
                "digestType": mime_type,
                "size": os.path.getsize(first_image_path),
                "mimeType": mime_type
            },
            "fileData": encoded_path,
            "type": {
                "uri": "/dk/atira/pure/equipment/equipmentfiles/photo",
                "term": {
                    "en_US": "Photo",
                    "es_CO": "Fotografía"
                }
            }
        }]
    }

    # Construir la URL para la solicitud PUT
    url = f"{base_url}/{row['UUID']}"

    # Enviar la solicitud PUT a la API
    response = requests.put(url, headers=headers, json=data)
    print(url)

    # Revisar la respuesta de la API
    if response.status_code == 200:
        print(f"La solicitud PUT para {row['Plaqueta']} fue exitosa.")
    else:
        print(f"Error en la solicitud PUT para {row['Plaqueta']}. Código de estado: {response.status_code}")
        print("Detalles:", response.text)

"""
Continuación de la explicación del script:

La parte final del script asegura que todas las imágenes correspondientes a los equipos listados en el archivo Excel 
se procesen y carguen correctamente en la API. Cualquier error o carga no exitosa se informa en la consola 
para facilitar la depuración.

Pasos clave:
------------
1. Para cada entrada de equipo en `imagenes_con_ruta`:
   - Extraer la ruta y el nombre del archivo de la imagen asociada.
   - Determinar el tipo MIME de la imagen según su extensión de archivo.
   - Leer y codificar la imagen en formato Base64.
   - Construir un payload JSON que contenga la imagen y los metadatos.
   - Enviar una solicitud PUT para actualizar la entrada del equipo con los datos de la imagen.
2. Para cada intento de carga:
   - Registrar en la consola el éxito o el fallo con detalles relevantes.

Manejo de errores:
------------------
- El script registra los errores encontrados durante las solicitudes a la API. 
  Esto incluye el código de estado HTTP y los detalles de la respuesta para ayudar a identificar problemas.

Personalización:
----------------
- Actualice `base_url` con el endpoint correcto de la API.
- Asegúrese de especificar la clave API correcta en los `headers`.
- Modifique las rutas de archivo y nombres de hojas para que coincidan con su configuración local.

Mejoras:
--------
1. Añadir un manejo robusto de excepciones para:
   - Errores de lectura/escritura de archivos.
   - Problemas de conectividad de red durante las llamadas a la API.
2. Implementar registro de logs para guardar el estado de las cargas en un archivo de registro en lugar de depender únicamente de la salida en consola.
3. Validar los datos del archivo Excel para asegurarse de que no haya campos faltantes o inválidos antes de procesarlos.

Fin de la documentación del script
"""