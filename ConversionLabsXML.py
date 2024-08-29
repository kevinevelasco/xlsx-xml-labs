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
#nlp = spacy.load("es_core_news_sm")
Language.factory("language_detector",func=get_lang_detector)
#nlp.add_pipe('language_detector',last=True)


#Create Element Tree root class
upmprojects = ET.Element("upmprojects")
upmprojects.set('xmlns',"v1.upmproject.pure.atira.dk")
upmprojects.set('xmlns:ns2',"v3.commons.pure.atira.dk")

#Upload the xlsx file for test
print("hola mundo")
