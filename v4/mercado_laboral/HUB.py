import pandas as pd
import numpy as np
import time
from HUB_DAMAC.v4.mercado_laboral.indicadores import indicadores_xarea_BR
from HUB_DAMAC.v4.mercado_laboral.scrap import scraping_BR,scraping_informalidad,scraping_desempleo_desetacionalizada_mensual,\
scraping_desempleo_sexo,scraping_desempleo_regiones,scraping_desempleo_estacionalizado
from HUB_DAMAC.v4.mercado_laboral.excel import guardar_excel
from HUB_DAMAC.v4.mercado_laboral.limpieza import clean_mlaboral_BR,clean_informalidad,clean_desempleo_desestacionalizado,\
clean_desempleo_empleo_sexo,clean_desempleo_empleo_regiones,clean_desempleo_estacionalizado


def actualizar_BR(carpeta,actualizar_todo = False,indicadores='',excel=False,hipervinculos=False,t=0):
    """
    actualizar_BR permite actualizar todos los indicadores o indicadores específicos que se encuentren 
    en la página del Banco de la República https://totoro.banrep.gov.co/estadisticas-economicas/ 
    
        ¡Tenga en cuenta que el código funciona únicamente con google chrome!
    
    si desea saber que cuáles son los indicadores disponibles para el mercado laboral utilice el móduolo indicadores_xarea_BR(0)
    
    Parámetros:
    -----------
    
    carpeta:
    tipo -> tring
            
             especifique la ruta de la carpeta en la que desea almacenar la información, por ejemplo:
             
             "C:/Escritorio/MercadoLaboral" 
             o
             r"C:\Escritorio\MercadoLaboral"
             
             Tenga en cuenta que debe utilizar correctamente alguna de las dos formas bien sea "C:/carpeta/subcarpeta" 
             o r"C:\carpeta\subcarpeta"
             
    actualizar_todo:
    tipo -> bool
            
            True: Actualiza todos los indicadores del mercado laboral presentes en la página del Banco de la República
            False: Actualiza solo los indicadores especificados en el parámetro indicadores
            
    indicadores:
    tipo -> string
    
            Sii actualizar_todo = False, entonces especifique en una lista los indicadores que desea actualizar, por ejemplo:
            
            indicadores = ['Tasa de desempleo','Tasa de ocupación']
            
    excel:
    tipo -> bool
    
            True: genera un excel que contiene todos los archivos descargados y la fuente de los mismos
            False: no genera el excel
            
    hipervinculos:
    tipo -> bool
    
            True: incluye hipervinculos en el excel
            False: no incluye hipervinculos en el excel
            
    t:
    tipo -> int
    
            añade un tiempo de espera para realizar el scraping web, esto es útil si la conexión de internet es inestable,
            o la descarga de los archivos no se completa correctamente
            
            si t = 10 agrega 10 segundos de espera en cada indicador 
            
    --------
    
    ejemplos:
    
    laboral = HUB.actualizar_BR(excel=True, hipervinculos=True, t=20,
                                        actualizar_todo=True, 
                                        carpeta=r"D:\Desktop\MercadoLaboral")
                                        
    [out] : genera un excel con todos los indicadores actualizados reportados en la página del banco de la república
    
    
    comentarios finales:
    
    desactive momentáneamente el firewall de windows o concedale permisos de adminsitrador a python para que pueda ejecutar correctamente
    el scraping web
    """
    indicadores_BR = indicadores_xarea_BR(0,tiempo=t)
    indicadores_BR = [i for i in indicadores_BR]
    
    fuente_BR = pd.DataFrame({})
    fuente_BR['Indicador'] = indicadores_BR
    fuente_BR['Frecuencia'] = 'Mensual'
    fuente_BR['Fuente'] = "https://totoro.banrep.gov.co/estadisticas-economicas/"
    
    if actualizar_todo:
        for i in fuente_BR['Indicador']:
            i = scraping_BR(0,indicador=i,path=carpeta+r"\BANREP",tiempo=t)
            i_clean = clean_mlaboral_BR(path=carpeta+r"\BANREP")
    else:
        for i in indicadores:
            i = scraping_BR(0,indicador=i,path=carpeta+r"\BANREP",tiempo=t)
            i_clean = clean_mlaboral_BR(path=carpeta+r"\BANREP")
    
    
    if excel:
        excel = guardar_excel(Fuente=fuente_BR,carpeta_origen=carpeta+r"\BANREP\CSV",carpeta_destino=carpeta,nombre_archivo='HUB_MercadoLaboral_BR',hyperlinks=hipervinculos)
        
            
    return ('Indicadores del Banco de la República Actualizados con éxtio')
    
def actualizar_DANE(carpeta,actualizar_todo = False,indicadores='',excel=False,hipervinculos=False,t=0):
    """
    actualizar_BR permite actualizar todos los indicadores o indicadores específicos usados por la DAMAC que se encuentren 
    en la página del Banco del DANE https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral
    
        ¡Tenga en cuenta que el código funciona únicamente con google chrome!
    
    indicadores usados:
    
                + Informalidad  
                + Desempleo_desestacionalizado 
                + Desempleo_por_sexo
                + Desempleo_por_region 
                + Desempleo_estacionalizado 
    
    Parámetros:
    -----------
    
    carpeta:
    tipo -> tring
            
             especifique la ruta de la carpeta en la que desea almacenar la información, por ejemplo:
             
             "C:/Escritorio/MercadoLaboral" 
             o
             r"C:\Escritorio\MercadoLaboral"
             
             Tenga en cuenta que debe utilizar correctamente alguna de las dos formas bien sea "C:/carpeta/subcarpeta" 
             o r"C:\carpeta\subcarpeta"
             
    actualizar_todo:
    tipo -> bool
            
            True: Actualiza todos los indicadores usados por la DAMAC del mercado laboral presentes en el DANE
            False: Actualiza solo los indicadores especificados en el parámetro indicadores
            
    indicadores:
    tipo -> string
    
            Sii actualizar_todo = False, entonces especifique en una lista los indicadores que desea actualizar, por ejemplo:
            
            indicadores = ['Tasa de desempleo','Tasa de ocupación']
            
    excel:
    tipo -> bool
    
            True: genera un excel que contiene todos los archivos descargados y la fuente de los mismos
            False: no genera el excel
            
    hipervinculos:
    tipo -> bool
    
            True: incluye hipervinculos en el excel
            False: no incluye hipervinculos en el excel
            
    t:
    tipo -> int
    
            añade un tiempo de espera para realizar el scraping web, esto es útil si la conexión de internet es inestable,
            o la descarga de los archivos no se completa correctamente
            
            si t = 10 agrega 10 segundos de espera en cada indicador 
            
    --------
    
    ejemplos:
    
    laboral = HUB.actualizar_BR(excel=True, hipervinculos=True, t=20,
                                        actualizar_todo=True, 
                                        carpeta=r"D:\Desktop\MercadoLaboral")
                                        
    [out] : genera un excel con todos los indicadores actualizados reportados en la página del banco de la república
    
    
    comentarios finales:
    
    desactive momentáneamente el firewall de windows o concedale permisos de adminsitrador a python para que pueda ejecutar correctamente
    el scraping web
    
    """
    

    fuente_DANE = pd.DataFrame({
        'Indicador':['Informalidad',
        'Desempleo_desestacionalizado',
        'Desempleo_por_sexo',
        'Desempleo_por_region',
        'Desempleo_estacionalizado'],

        'Frecuencia': ['Trimestral','Mensual','Trimestre móvil','Semestral','Trimestre móvil'],

        'Fuente': ['https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-informal-y-seguridad-social',
        'https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-y-desempleo/mercado-laboral-historicos',
        "https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/segun-sexo/mercado-laboral-historicos",
        'https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/por-regiones/mercado-laboral-por-regiones-historicos',
        'https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-y-desempleo/mercado-laboral-historicos']
        })
    

    if actualizar_todo:
        for i in fuente_DANE['Indicador']:
            if i == 'Informalidad':
                informalidad = scraping_informalidad(path=carpeta+r"\DANE",tiempo=t)
                clean_infor = clean_informalidad(path=carpeta+r"\DANE")
                
            elif i == 'Desempleo_desestacionalizado':
                desestacionalizado = scraping_desempleo_desetacionalizada_mensual(path=carpeta+r"\DANE",tiempo=t)
                clean_deses = clean_desempleo_desestacionalizado(path=carpeta+r"\DANE")
                
            elif i == 'Desempleo_por_sexo':
                sexo = scraping_desempleo_sexo(path=carpeta+r"\DANE",tiempo=t)
                clean_sexo = clean_desempleo_empleo_sexo(path=carpeta+r"\DANE")
                
            elif i == 'Desempleo_por_region':
                region = scraping_desempleo_regiones(path=carpeta+r"\DANE",tiempo=t)
                clean_region = clean_desempleo_empleo_regiones(path=carpeta+r"\DANE")
                
            elif i == 'Desempleo_estacionalizado':
                estacionalizado = scraping_desempleo_estacionalizado(path=carpeta+r"\DANE",tiempo=t)
                clean_estac = clean_desempleo_estacionalizado(path=carpeta+r"\DANE")
    else:
        for i in indicadores:
            if i == 'Informalidad':
                informalidad = scraping_informalidad(path=carpeta+r"\DANE",tiempo=t)
                clean_infor = clean_informalidad(path=carpeta+r"\DANE")
                
            elif i == 'Desempleo_desestacionalizado':
                desestacionalizado = scraping_desempleo_desetacionalizada_mensual(path=carpeta+r"\DANE",tiempo=t)
                clean_deses = clean_desempleo_desestacionalizado(path=carpeta+r"\DANE")
                
            elif i == 'Desempleo_por_sexo':
                sexo = scraping_desempleo_sexo(path=carpeta+r"\DANE",tiempo=t)
                clean_sexo = clean_desempleo_empleo_sexo(path=carpeta+r"\DANE")
                
            elif i == 'Desempleo_por_region':
                region = scraping_desempleo_regiones(path=carpeta+r"\DANE",tiempo=t)
                clean_region = clean_desempleo_empleo_regiones(path=carpeta+r"\DANE")
                
            elif i == 'Desempleo_estacionalizado':
                estacionalizado = scraping_desempleo_estacionalizado(path=carpeta+r"\DANE",tiempo=t)
                clean_estac = clean_desempleo_estacionalizado(path=carpeta+r"\DANE")
                
            else:
                print('Indicador no válido, verifique que esté escrito correctamente')

    if excel:
        excel = guardar_excel(Fuente=fuente_DANE,carpeta_origen=carpeta+r"\DANE\CSV",carpeta_destino=carpeta,nombre_archivo='HUB_MercadoLaboral_DANE',hyperlinks=hipervinculos)

    

                
    


