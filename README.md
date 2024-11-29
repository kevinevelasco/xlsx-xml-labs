<div align="center">

# **Pontificia Universidad Javeriana**  
## *Monitoría Vicerrectoría de la Investigación*  
### Monitor: Kevin Estiven Velasco Pinto  

---

</div>
Este repositorio contiene un conjunto de herramientas y scripts diseñados para realizar diversas operaciones de conversión y carga de datos relacionados con laboratorios y equipos. A continuación, se describe la estructura general del proyecto y cómo usar cada componente.

---

## Estructura del Proyecto

El repositorio está organizado en las siguientes carpetas y archivos:

### 1. **APIRequestCargueImagenesEquipos**
- Contiene un script en Python que realiza peticiones PUT para cargar imágenes de equipos a una API (en este caso, la API de Pure).
- Incluye:
  - Un archivo Python con el código para realizar la petición PUT.
  - Un archivo JSON con el formato del payload necesario para las solicitudes.
  - Dos archivos Excel con información de los equipos.
  
### 2. **ConversionExcelXMLEquipos**
- Contiene un script en Python que convierte datos desde un archivo Excel a un formato XML, específicamente para información relacionada con equipos.

### 3. **ConversionExcelXMLLabs**
- Similar a la carpeta anterior, pero orientada a la conversión de datos de laboratorios desde Excel a XML.

### 4. **ConversionPDFaExcel**
- Contiene un script en Python que convierte documentos en formato PDF a Excel, facilitando la extracción y manipulación de datos tabulares.

### 5. Archivos adicionales:
- **`.gitignore`**: Define los archivos y carpetas que deben ser ignorados por Git.
- **`requirements.txt`**: Lista de dependencias de Python necesarias para ejecutar los scripts.

---

## Requisitos Previos

Se recomienda crear un ambiente virtual para gestionar las dependencias y evitar conflictos con otros proyectos.

### Pasos para configurar el ambiente virtual:
1. Crear el ambiente virtual:
   ```bash
   python -m venv env
2. Activar el ambiente virtual:
- En Windows:
     ```bash
   .\env\Scripts\activate
- En macOS/Linux:
     ```bash
   source env/bin/activate
3. Instalar las dependencias:
     ```bash
   pip install -r requirements.txt
---

## Cómo Ejecutar los Scripts
### Carga de Imágenes a la API
1. Navegar a la carpeta APIRequestCargueImagenesEquipos.
2. Asegurarse de tener el archivo JSON con el payload y los Excel con la información de los equipos.
3. Ejecutar el script de carga:
   
     ```bash
    python ImgAPIRequest.py
### Conversión de Excel a XML
#### Para equipos:
1. Navegar a la carpeta ConversionExcelXMLEquipos.
2. Ejecutar el script:
     ```bash
    python ConversionEquiposXML.py
####  Para laboratorios:
1. Navegar a la carpeta ConversionExcelXMLLabs.
2. Ejecutar el script:
   
     ```bash
    python ConversionLabsXML.py
### Conversión de PDF a Excel
1. Navegar a la carpeta ConversionPDFaExcel.
2. Asegurarse de tener los PDFs en la carpeta adecuada.
3. Ejecutar el script:
   
     ```bash
    python PDFConversion.py
