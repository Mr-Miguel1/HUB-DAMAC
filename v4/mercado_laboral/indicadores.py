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

#### Verificar el correcto funcionamiento ###
def indicadores_xarea_BR(area,tiempo=0):
    """
    Esta funcion devuelve una lista con todos los indicadores por área reportados en la página del banco de la república
    https://totoro.banrep.gov.co/estadisticas-economicas/
    
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
            area = 9---> Metales Preciosos
            area = 10---> Unidad de Valor Real, UVR
            area = 11---> Reservas Internacionales
            area = 12---> Sector Fiscal
            area = 13---> Índice de tasa de cambio real base 2010
            area = 14---> Sector Externo
            area = 15---> Tasas de cambio nominales
            area = 16---> Agregados monetarios
            area = 17---> Agregados crediticios
            area = 18---> Tasa de interés
            
    NOTA: Se recomienda tener el cortafuegos desactivado o darle permisos de administrador para tener acceso a la red
          pública
    """
    # opciones de navegación
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-extensions')

    #Excepciones
    ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,)
    
    ###########################################################
    ################ OJO #####################################

    # Deben descargar el chromedirver.exe en https://chromedriver.chromium.org/ si su navegador prederminado es chrome
    driver_path = "C:/Users/Laptop/HUB_DAMAC/chromedriver.exe" #corresponde a la carpeta en la que guardo el archivo
    driver = webdriver.Chrome(driver_path, options=options,)
    
    ##########################################################
    #########################################################

    # iniciamos la página web
    driver.get("https://totoro.banrep.gov.co/estadisticas-economicas/")

    # Damos click en el catálogo de series
    WebDriverWait(driver,5+tiempo)    .until(EC.element_to_be_clickable((By.ID,"idBtnCatalogoSeries")))    .click()
    
    l1 = pd.Series(driver.find_element_by_id("maincontent:datacatalogopadre:{}:datacatalogohijo_data".format(str(area))).text).str.split("\n")[0]
    l2 = pd.Series(l1).str.split(',').index.to_list()
    
    indicadores = []
    for n in l2:
        try:
            serie = driver.find_element_by_xpath("//tbody[@id='maincontent:datacatalogopadre:{}:datacatalogohijo_data']/tr[@data-ri='{}']/td[@class='columnDE']".format(str(area),str(n)))
            indicadores.append(serie.text)
        except:
            continue
    driver.close()
    return pd.Series(indicadores)

