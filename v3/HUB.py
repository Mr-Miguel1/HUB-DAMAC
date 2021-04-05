import pandas as pd
import numpy as np
import time
from indicadores import indicadores_xarea
from scrap import scraping_BR,scraping_DANE_mercado_laboral
from excel import guardar_excel
from limpieza import clean_mlaboral_BR,clean_informalidad,clean_desempleo_desestacionalizado,\
clean_desempleo_empleo_sexo,clean_desempleo_empleo_regiones,clean_desempleo_estacionalizado

    
def actualizar_mercado_laboral(carpeta,actualizar_todo = False,indicadores='',excel=False,hipervinculos=False,t=0):

    fuente_laboral = pd.DataFrame({
        'Indicador':['Tasa de desempleo',
        'Tasa de ocupación',
        'Tasa global de participación',
        'Informalidad',
        'Desempleo_desestacionalizado',
        'Desempleo_por_sexo',
        'Desempleo_por_region',
        'Desempleo_estacionalizado'],

        'Frecuencia': ['Mensual','Mensual','Mensual','Trimestral','Mensual','Trimestre móvil','Semestral','Trimestre móvil'],

        'Fuente': ['https://totoro.banrep.gov.co/estadisticas-economicas/',
        'https://totoro.banrep.gov.co/estadisticas-economicas/',
        'https://totoro.banrep.gov.co/estadisticas-economicas/',
        'https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-informal-y-seguridad-social',
        'https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-y-desempleo/mercado-laboral-historicos',
        "https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/segun-sexo/mercado-laboral-historicos",
        'https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/por-regiones/mercado-laboral-por-regiones-historicos',
        'https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-y-desempleo/mercado-laboral-historicos']
        })

    lista_indicadores_BR = ['Tasa de desempleo','Tasa de ocupación','Tasa global de participación']

    if actualizar_todo:
        for i in fuente_laboral['Indicador']:
            if i in lista_indicadores_BR:
                try:
                    i = scraping_BR(0,indicador=i,path=carpeta,tiempo=t)
                except:
                    continue
            elif i == 'Informalidad':
                informalidad = scraping_DANE_mercado_laboral().informalidad(path=carpeta,tiempo=t)
            elif i == 'Desempleo_desestacionalizado':
                desestacionalizado = scraping_DANE_mercado_laboral().desempleo_desetacionalizada_mensual(path=carpeta,tiempo=t)
            elif i == 'Desempleo_por_sexo':
                sexo = scraping_DANE_mercado_laboral().desempleo_sexo(path=carpeta,tiempo=t)
            elif i == 'Desempleo_por_region':
                region = scraping_DANE_mercado_laboral().desempleo_regiones(path=carpeta,tiempo=t)
            elif i == 'Desempleo_estacionalizado':
                estacionalizado = scraping_DANE_mercado_laboral().desempleo_estacionalizado(path=carpeta,tiempo=t)
    else:
        for i in indicadores:
            if i in lista_indicadores_BR:
                try:
                    i = scraping_BR(0,indicador=i,path=carpeta,tiempo=t)
                except:
                    continue
            elif i == 'Informalidad':
                informalidad = scraping_DANE_mercado_laboral().informalidad(path=carpeta,tiempo=t)
            elif i == 'Desempleo_desestacionalizado':
                desestacionalizado = scraping_DANE_mercado_laboral().desempleo_desetacionalizada_mensual(path=carpeta,tiempo=t)
            elif i == 'Desempleo_por_sexo':
                sexo = scraping_DANE_mercado_laboral().desempleo_sexo(path=carpeta,tiempo=t)
            elif i == 'Desempleo_por_region':
                region = scraping_DANE_mercado_laboral().desempleo_regiones(path=carpeta,tiempo=t)
            elif i == 'Desempleo_estacionalizado':
                estacionalizado = scraping_DANE_mercado_laboral().desempleo_estacionalizado(path=carpeta,tiempo=t)
            else:
                print('Indicador no válido, verifique que esté escrito correctamente')

    try:
        clean_mlaboral_BR(path=carpeta)
    except:
        pass        
    try:
        clean_informalidad(path=carpeta)
    except:
        pass
    try:
        clean_desempleo_desestacionalizado(path=carpeta)
    except:
        pass
    try:
        clean_desempleo_empleo_sexo(path=carpeta)
    except:
        pass
    try:
        clean_desempleo_empleo_regiones(path=carpeta)
    except:
        pass
    try:
        clean_desempleo_estacionalizado(path=carpeta)
    except:
        pass


    if excel:
        excel = guardar_excel(Fuente=fuente_laboral,carpeta_origen=carpeta,carpeta_destino=carpeta,nombre_archivo='hub_laboral',hyperlinks=hipervinculos)
    else:
        print('Esta bien no genero el excel')

    

                
    


