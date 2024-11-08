from requests.auth import HTTPBasicAuth
import requests
import pandas as pd
import os


url = 'https://puj-staging.elsevierpure.com/ws/api/equipment'
headers = {
    'Accept': 'application/json',
    'api-key': ''
}

# req = requests.get(url, headers=headers)
# print(req.json())

# Read excel and convert to dataframe
dataframe1 = pd.read_excel(
    r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Excel\Formato Infraestructura - Perfiles y Capacidades - Ajustado.xlsx', sheet_name='uuids', dtype=object)
column = dataframe1.UUID.unique()

image_folder_path = r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Fotos'
image_formats = ['.png', '.jpg', '.jpeg', '.gif']

equipos = []
for uid in column:
      #Create Element Tree root class
    temp_eq = dataframe1[dataframe1["UUID"]==uid]
    temp_eq_id = str(uid)
    temp_eq_source_id = temp_eq["Source ID"].values[0]
    temp_eq_plaqueta = temp_eq["Plaqueta"].values[0]
    temp_eq_nombre = temp_eq["Nombre Equipo"].values[0]

    # Buscar la imagen usando el valor de "Plaqueta"
    temp_eq_image_path = None
    for fmt in image_formats:
        image_path = os.path.join(image_folder_path, f"{temp_eq_plaqueta}{fmt}")
        if os.path.exists(image_path):
            temp_eq_image_path = os.path.basename(image_path)  # Guarda la ruta si la imagen existe
            break  # Deja de buscar una vez encontrada la primera imagen

    # Añadir los datos del equipo a la lista
    equipos.append({
        "Nombre Equipo": temp_eq_nombre,
        "UUID": temp_eq_id,
        "Source ID": temp_eq_source_id,
        "Plaqueta": temp_eq_plaqueta,
        "Image Path": temp_eq_image_path  # Ruta de la imagen o None
    })


equipos_df = pd.DataFrame(equipos)
# print(equipos_df)

imagenes_con_ruta = equipos_df[equipos_df['Image Path'].notna()]

# Dejar solo el nombre del archivo con su extensión en el campo 'Image Path'
imagenes_con_ruta['Image Path'] = imagenes_con_ruta['Image Path'].apply(os.path.basename)

# Contar cuántas imágenes tienen una ruta válida
cantidad_imagenes_con_ruta = imagenes_con_ruta.shape[0]

# Mostrar los resultados
# print(f"Cantidad de imágenes con ruta: {cantidad_imagenes_con_ruta}")
# print("Imágenes con ruta:")
# print(imagenes_con_ruta[['UUID', 'Source ID', 'Plaqueta', 'Image Path']])

print(equipos_df)

# Crear un archivo Excel a partir del DataFrame equipos_df
equipos_df.to_excel(
    r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Archivo Python\xlsx-xml-labs\formatoJSON-photos-equipo\equipos_data.xlsx',  # Reemplaza con la ruta y el nombre que desees
    index=False,  # No incluir el índice del DataFrame en el archivo Excel
    sheet_name='Equipos'  # Nombre de la hoja dentro del archivo Excel
)

print("Archivo Excel creado exitosamente.")