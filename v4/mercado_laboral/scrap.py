import pandas as pd
import numpy as np

from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
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

#### Por favor descargue el webdrvier de Chrome disponible en https://chromedriver.chromium.org/ y extragiga el archivo y especifique
#### la carpeta en la que esta almacenado, en mi caso es:

carpeta_driver = "C:/Users/Laptop/HUB_DAMAC/chromedriver.exe" ###coloque aquí la ruta de su carpeta 

def scraping_BR(area, indicador,
                path, tiempo):
    """
    Esta función permite scrapear la página web del banco de la república que contiene la mayoría de indicadores
    más importantes para la economía Colombiana, página: https://totoro.banrep.gov.co/estadisticas-economicas/ 
    
    
    De tal forma que ingresa y descarga un archivo csv con el indicador solicitado del área correspondiente 
    
    Parámetros:
    ----------------
    area : corresponde al área que dispone el banco de la república para la organización de los indicadores
    
            area = 0 ---> Mercado Laboral
            area = 1 ---> Producto Interno Bruto, base 2015
            area = 2 ---> Producto Interno Bruto, base 2005
            area = 3 ---> Producto Interno Bruto, base 2000
            area = 4 ---> Producto Interno bruto, base 1994
            area = 5 ---> Precios e inflación, base 2014
            area = 6 ---> Índices de precios de vivienda, base 1990
            area = 7 ---> Inflación total y meta
            area = 8 ---> Índices de precios al consumidor, base 2018
            area = 9 ---> Metales Preciosos
            area = 10---> Unidad de Valor Real, UVR
            area = 11---> Reservas Internacionales
            area = 12---> Sector Fiscal
            area = 13---> Índice de tasa de cambio real base 2010
            area = 14---> Sector Externo
            area = 15---> Tasas de cambio nominales
            area = 16---> Agregados monetarios
            area = 17---> Agregados crediticios
            area = 18---> Tasa de interés
            
    indicador: corresponde al indicador específico en cada area, si desea ver cual es la lista de indicadores 
                para cada área use la función indicadores_xarea()
                
    path: selecciona la carpeta en la que se descarga el archivo
    
    NOTA: Se recomienda tener el cortafuegos desactivado o darle permisos de administrador para tener acceso a la red
          pública
    """
    
    ### configuramos las opciones de inicio del navegador experimental
    
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-extensions')
    options.add_experimental_option("prefs", {"download.default_directory":path, "download.prompt_for_download": False,"download.directory_upgrade": True,"safebrowsing.enabled": True}) 
    
    #Excepciones
    ignored_exceptions = (NoSuchElementException,StaleElementReferenceException,)
    
    ###########################################################
    ################ OJO #####################################

    # Deben descargar el chromedirver.exe en https://chromedriver.chromium.org/ si su navegador prederminado es chrome
    driver_path = carpeta_driver #corresponde a la carpeta en la que guardo el archivo
    driver3 = webdriver.Chrome(driver_path, options=options,)
    
    ##########################################################
    #########################################################
    
    # iniciamos la página web
    
    driver3.get("https://totoro.banrep.gov.co/estadisticas-economicas/")

    # damos click en el catálogo de series
    WebDriverWait(driver3,5+tiempo)\
    .until(EC.element_to_be_clickable((By.ID,"idBtnCatalogoSeries")))\
    .click()
    
    # Obtenemos la longitud de cada catálogo
    
    l1 = pd.Series(driver3.find_element_by_id("maincontent:datacatalogopadre:{}:datacatalogohijo_data".format(str(area))).text).str.split("\n")[0]
    l2 = pd.Series(l1).str.split(',').index.to_list()

    ## Iteramos sobre el catálogo respectivo y damos click en el botón de descarga si el indicador que estamos buscando
    ## esta presente en la página web
    
    for n in l2:
        try:
            serie = driver3.find_element_by_xpath("//tbody[@id='maincontent:datacatalogopadre:{}:datacatalogohijo_data']/tr[@data-ri='{}']/td[@class='columnDE']".format(str(area),str(n)))
            if serie.text == indicador:
                
                WebDriverWait(driver3,5+tiempo,ignored_exceptions=ignored_exceptions)\
                .until(EC.element_to_be_clickable((By.ID,"maincontent:datacatalogopadre:{}:datacatalogohijo:{}:indicadorId".format(str(area),str(n)))))\
                .click()

                time.sleep(5+tiempo) #Agregamos un tiempo de espera para que se cargue la página completamente
                driver3.find_element_by_class_name("highcharts-button".replace(' ','.')).click()
                s = driver3.find_element_by_class_name("highcharts-contextmenu")
                h = s.find_element_by_xpath("//ul[@class='highcharts-menu']/li[2]")
                h.click() 
        except:
            pass
            
    
    time.sleep(5+tiempo)
    # Renombramos el archivo con el nombre del indicador
    try:
        shutil.move(path+"\chart.csv",path+"\{}.csv".format(indicador.replace(':',',')))
    except:
        print('Asegurese de cerrar excel')
        pass
    
    driver3.close()
    return('la descarga de {} fue exitosa'.format(indicador))


    
def scraping_informalidad(path,tiempo):

    """ Path: Es la carpeta donde se guarda el archivo  """

    ### configuramos las opciones de inicio del navegador experimental

    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-extensions')
#         options.add_argument('--headless')
    options.add_experimental_option("prefs", {"download.default_directory":path, "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True}) 

    #Excepciones
    ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,)

    ###########################################################
    ################ OJO #####################################

    # Deben descargar el chromedirver.exe en https://chromedriver.chromium.org/ si su navegador prederminado es chrome
    driver_path = carpeta_driver #corresponde a la carpeta en la que guardo el archivo
    driver = webdriver.Chrome(driver_path, options=options)

    ##########################################################
    #########################################################

    # iniciamos la página web
    driver.get("https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-informal-y-seguridad-social")

    try:
        WebDriverWait(driver,5+tiempo)\
        .until(EC.element_to_be_clickable((By.LINK_TEXT,"Anexos")))\
        .click()
    except:
        print('Hubo un problema al descargar la tasa de informalidad del dane')

    time.sleep(5+tiempo)
    driver.close()
    return ('La descarga de la tasa de informalidad fue exitosa')

def scraping_desempleo_desetacionalizada_mensual(path,tiempo):
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-extensions')
    options.add_experimental_option("prefs", {"download.default_directory":path, "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True}) 

    #Excepciones
    ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,)

    ###########################################################
    ################ OJO #####################################

    # Deben descargar el chromedirver.exe en https://chromedriver.chromium.org/ si su navegador prederminado es chrome
    driver_path = carpeta_driver #corresponde a la carpeta en la que guardo el archivo
    driver = webdriver.Chrome(driver_path, options=options)

    driver.get("https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-y-desempleo#geih-mercado-laboral")

    full_xpath = "/html/body/div[1]/div[2]/div/div/div/article/section/div[3]/div/div[1]/table/tbody/tr/td/div/table/tbody/tr[3]/td[4]/a"
    try:
        WebDriverWait(driver,5+tiempo)\
        .until(EC.element_to_be_clickable((By.XPATH,full_xpath)))\
        .click()
    except:
        print('Hubo un problema al descargar el desempleo desestacionalizado mensual')

    time.sleep(5+tiempo)
    driver.close()
    return ('La descarga de la tasa desempleo desestacionalizado mensual fue exitosa')


def scraping_desempleo_sexo(path,tiempo):
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-extensions')
    options.add_experimental_option("prefs", {"download.default_directory":path, "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True}) 

    #Excepciones
    ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,)

    ###########################################################
    ################ OJO #####################################

    # Deben descargar el chromedirver.exe en https://chromedriver.chromium.org/ si su navegador prederminado es chrome
    driver_path = carpeta_driver #corresponde a la carpeta en la que guardo el archivo
    driver = webdriver.Chrome(driver_path, options=options)

    driver.get("https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/segun-sexo/mercado-laboral-historicos")

    try:
        WebDriverWait(driver,5+tiempo)\
        .until(EC.element_to_be_clickable((By.LINK_TEXT,"Anexos")))\
        .click()
    except:
        print('Hubo un problema al descargar el desempleo por sexo')    
    time.sleep(15+tiempo)
    driver.close()
    return ('La descarga de la tasa de desempleo  por sexo fue exitosa')

def scraping_desempleo_regiones(path,tiempo):
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-extensions')
    options.add_experimental_option("prefs", {"download.default_directory":path, "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True}) 

    #Excepciones
    ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,)

    ###########################################################
    ################ OJO #####################################

    # Deben descargar el chromedirver.exe en https://chromedriver.chromium.org/ si su navegador prederminado es chrome
    driver_path = carpeta_driver #corresponde a la carpeta en la que guardo el archivo
    driver = webdriver.Chrome(driver_path, options=options)

    driver.get("https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/por-regiones/mercado-laboral-por-regiones-historicos")

    try:
        WebDriverWait(driver,5+tiempo)\
        .until(EC.element_to_be_clickable((By.LINK_TEXT,"Anexos")))\
        .click()
    except:
        print('Hubo un problema al descargar el desempleo por regiones')    
    time.sleep(15+tiempo)
    driver.close()
    return ('La descarga de la tasa de desmepleo por por regiones fue exitosa')

def scraping_desempleo_estacionalizado(path,tiempo):
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-extensions')
    options.add_experimental_option("prefs", {"download.default_directory":path, "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True}) 

    #Excepciones
    ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,)

    ###########################################################
    ################ OJO #####################################

    # Deben descargar el chromedirver.exe en https://chromedriver.chromium.org/ si su navegador prederminado es chrome
    driver_path = carpeta_driver #corresponde a la carpeta en la que guardo el archivo
    driver = webdriver.Chrome(driver_path, options=options)

    driver.get("https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-y-desempleo#geih-mercado-laboral")
    full_xpath = "/html/body/div[1]/div[2]/div/div/div/article/section/div[3]/div/div[1]/table/tbody/tr/td/div/table/tbody/tr[2]/td[4]/a"

    try:
        WebDriverWait(driver,5+tiempo)\
        .until(EC.element_to_be_clickable((By.XPATH,full_xpath)))\
        .click()
    except:
        print('Hubo un problema al descargar el desempleo estacionalizado')

    time.sleep(5+tiempo)
    driver.close()
    return ('La descarga de la tasa desempleo estacioanlizado fue exitosa')