#!/usr/bin/env python
# coding: utf-8

# In[3]:


import pandas as pd
import numpy as np

from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import WebDriverException

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

import urllib3
import csv

import time
import os
import shutil
import sys

def iniciador(path_descarga,
             url,
             full_xpath,
             indicador,
             tiempo): 
    """
    iniciador_DANE(path_descarga,driver_path) permite iniciar pestañas de chrome usando selenium
    
    Parámetros
    ---------
    
        path_descarga: carpeta para descargar la información
    """
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-extensions')
    options.add_experimental_option("prefs", {"download.default_directory":path_descarga, "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True}) 

    #Excepciones
    ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,)

    ###########################################################
    ################ OJO #####################################

    driver_path = r"C:\Users\Laptop\HUB_DAMAC\chromedriver.exe"
    try:
        driver = webdriver.Chrome(driver_path, options=options)
    except WebDriverException as we:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print("""
              ----
              Error
              ----
              
              ubicación: HUB_DAMAC\mercado_laboral
              script: scrapy.py
              función: iniciador
              linea de código: {}
              
              -----
              Información del error:
              -----
              
              + {}
              
              + Debe usar Google Chrome para poder ejecutar el código
              + Descargue el driver de chrome 'chromedirver.exe' en https://chromedriver.chromium.org/'
              + Extraiga el archivo en una carpeta
              + copie y pegue la ruta de la carpeta en driver_path
              """.format(exc_tb.tb_lineno,we))
              
    
    try:
        #Iniciamos el nevegador
        driver.get(url)

        # Inspeccionar el código de la página y encontrar el selector adecuado
        data_base_download = WebDriverWait(driver,5+tiempo)                            .until(EC.element_to_be_clickable((By.XPATH,full_xpath)))
        
        # Nombre del archivo
        href = data_base_download.get_attribute('href')
        file_name = href.split("/")[-1]
        
        # Clickeamos para descargar la base de datos
        data_base_download.click()
        
        # Agregamos Tiempo de espera para que descarge correctamente
        time.sleep(5+tiempo)
        driver.close()
        
        # Mensaje
        descargas = os.listdir(path_descarga)        
    
        if file_name not in descargas:
            print("\n",'Descarga no completada porfavor agregue más tiempo')
        else:
            print("""{}
            Descarga exitosa
            url : {}
            """.format(file_name, url))
        
    except TimeoutException as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(
             """
              ----
              Error
              ----
              
              ubicación: HUB_DAMAC\mercado_laboral
              script: scrap.py
              función: {}
              linea del código: {}
              
              tipo: {}
              
              
              -----
              Información del error:
              -----
              + {} 
              """.format(indicador,exc_tb.tb_lineno,exc_type,ex))

        driver.close()
        pass

