import pandas as pd
from spacy.language import Language
from spacy_langdetect import LanguageDetector
import spacy
import xml.etree.ElementTree as ET
import datetime

# Define language detector 
def get_lang_detector(nlp, name):
    return LanguageDetector()
#Excecute Natural Language processing in spanish to detect language
nlp = spacy.load("es_core_news_sm")
Language.factory("language_detector",func=get_lang_detector)
nlp.add_pipe('language_detector',last=True)


#Create Element Tree root class
result = ET.Element("result")
result.set('xmlns:xsi',"http://www.w3.org/2001/XMLSchema-instance")
result.set('xsi:schemaLocation',"http://localhost:8080/ws/api/524/xsd/schema1.xsd")

#Read excel and convert to dataframe
dataframe1 = pd.read_excel(r'C:\Users\DELL\OneDrive - Pontificia Universidad Javeriana\Gestión Monitores Perfiles\Infraestructura\Excel\Formato Infraestructura - Perfiles y Capacidades.xlsx', sheet_name='Laboratorios', dtype=object)

#print(dataframe1.head())

column = dataframe1.id_unico.unique()

for uid in column:
    #Create Element Tree root class
    temp_lab = dataframe1[dataframe1["id_unico"]==uid]
    temp_lab_id = str(uid)
    temp_lab_type = temp_lab["Tipo"].values[0]
    temp_lab_eng_name = temp_lab["Nombre Inglés"].values[0]
    temp_lab_es_name = temp_lab["Nombre español"].values[0]
    temp_lab_eng_description = temp_lab["Descripción Inglés"].values[0]
    temp_lab_es_description = temp_lab["Descripción Español"].values[0]
    temp_lab_address = temp_lab["Dirección"].values[0]
    temp_lab_phone = temp_lab["Número de contacto"].values[0]
    temp_lab_url = temp_lab["URL"].values[0]
    temp_lab_person = temp_lab["Nombre encargado"].values[0]
    temp_lab_created = temp_lab["Fecha creación Laboratorios"].values[0]
    temp_lab_loan_type = temp_lab["Disponible para préstamo"].values[0]
    temp_lab_keywords = temp_lab["Palabras clave"].values[0]
    temp_lab_services = temp_lab["Servicios del Laboratorio"].values[0]
    temp_lab_certifications = temp_lab["Certificaciones"].values[0]

    #print the first 5 rows of the dataframe

print(temp_lab.head())


