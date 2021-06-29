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

from HUB_DAMAC.v5.iniciador_scrap import iniciador

import urllib3
import csv

import time
import os
import shutil

#### Por favor descargue el webdrvier de Chrome disponible en https://chromedriver.chromium.org/ y extragiga el archivo y especifique
#### la carpeta en la que esta almacenado, en mi caso es:

carpeta_driver = "C:/Users/Laptop/HUB_DAMAC/chromedriver.exe" ###coloque aquí la ruta de su carpeta 

    
def scraping_informalidad(path,tiempo):
    
    url_ = "https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-informal-y-seguridad-social"
    full_xpath_ = "/html/body/div[1]/div[2]/div/div[1]/div/article/section/table[1]/tbody/tr/td/table/tbody/tr/td[2]/div/a"
    

    iniciador(path_descarga=path,
             url=url_,
             full_xpath=full_xpath_,
             indicador = "scraping_informalidad",
             tiempo=tiempo)

def scraping_desempleo_desetacionalizada_mensual(path,tiempo):
    
    url_ = "https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-y-desempleo"
    full_xpath_ = "/html/body/div[1]/div[2]/div/div/div/article/section/div[3]/div/div[1]/table/tbody/tr/td/div/table/tbody/tr[3]/td[4]/a"
    

    iniciador(path_descarga=path,
             url=url_,
             full_xpath=full_xpath_,
             indicador = "scraping_desempleo_desetacionalizada_mensual",
             tiempo=tiempo)


def scraping_desempleo_sexo(path,tiempo):
 
    url_ = "https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/segun-sexo"
    full_xpath_ = "/html/body/div[1]/div[2]/div/div[1]/div/article/section/table[1]/tbody/tr/td/table/tbody/tr/td[2]/div/a"
    

    iniciador(path_descarga=path,
             url=url_,
             full_xpath=full_xpath_,
             indicador = "scraping_desempleo_sexo",
             tiempo=tiempo)

def scraping_desempleo_regiones(path,tiempo):
    
    url_ = "https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/por-regiones"
    full_xpath_ = "/html/body/div[1]/div[2]/div/div[1]/div/article/section/table[1]/tbody/tr/td/table/tbody/tr/td[2]/div/a"
    

    iniciador(path_descarga=path,
             url=url_,
             full_xpath=full_xpath_,
             indicador = "scraping_desempleo_regiones",
             tiempo=tiempo)

def scraping_desempleo_estacionalizado(path,tiempo):
    
    url_ = "https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-y-desempleo"
    full_xpath_ = "/html/body/div[1]/div[2]/div/div/div/article/section/div[3]/div/div[1]/table/tbody/tr/td/div/table/tbody/tr[2]/td[4]/a"
    

    iniciador(path_descarga=path,
             url=url_,
             full_xpath=full_xpath_,
             indicador = "scraping_desempleo_estacionalizado",
             tiempo=tiempo)