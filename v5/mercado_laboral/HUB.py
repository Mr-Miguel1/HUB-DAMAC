import pandas as pd
import numpy as np
import time

from HUB_DAMAC.v5.iniciador_scrap import iniciador

from HUB_DAMAC.v5.mercado_laboral.scrap import scraping_informalidad,scraping_desempleo_desetacionalizada_mensual,\
scraping_desempleo_sexo,scraping_desempleo_regiones,scraping_desempleo_estacionalizado

from HUB_DAMAC.v5.excel import guardar_excel

from HUB_DAMAC.v5.mercado_laboral.limpieza_informalidad import clean_informalidad
from HUB_DAMAC.v5.mercado_laboral.limpieza_regiones import clean_regiones
from HUB_DAMAC.v5.mercado_laboral.limpieza_sexo import clean_sexo
from HUB_DAMAC.v5.mercado_laboral.limpieza_anexos_desestacionalizado import clean_anexos_desestacionalizado
from HUB_DAMAC.v5.mercado_laboral.limpieza_anexos_estacionalizado import clean_anexos_estacionalizado

from HUB_DAMAC.v5.mercado_laboral.fuente_ml import fuente_mercadolaboral


    
def actualizar_DANE(carpeta,
                    actualizar_todo = False,
                    indicadores='',
                    excel=False,
                    t=0):
    """
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
    tipo -> string
            
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
            
            
    t:
    tipo -> int
    
            añade un tiempo de espera para realizar el scraping web, esto es útil si la conexión de internet es inestable,
            o la descarga de los archivos no se completa correctamente
            
            si t = 10 agrega 10 segundos de espera en cada indicador 
            
    --------
    
    ejemplos:
    
    laboral = HUB.actualizar_BR(excel=True,
                                t=20,
                                actualizar_todo=True, 
                                carpeta=r"D:\Desktop\MercadoLaboral")
                                        
    [out] : genera un excel con todos los indicadores actualizados
    
    
    comentarios finales:
    
    desactive momentáneamente el firewall de windows o concedale permisos de adminsitrador a python para que pueda ejecutar correctamente
    el scraping web
    
    creador:
                Miguel Angel Manrique Rodriguez
                Estudiante de economía 2021-1S UNAL sede Bogotá
                mamanriquer@unal.edu.co
    
                :D por el derecho al acceso a la información open source
    """
    

    fuente_DANE = fuente_mercadolaboral()
    

    if actualizar_todo:
        
        # Informalidad
        informalidad = scraping_informalidad(path=carpeta+r"\DANE",tiempo=t)
        c_infor = clean_informalidad(path=carpeta+r"\DANE")
        
        # Anexos Desestacionalizados mensual
        desestacionalizado = scraping_desempleo_desetacionalizada_mensual(path=carpeta+r"\DANE",tiempo=t)
        c_deses = clean_anexos_desestacionalizado(path=carpeta+r"\DANE")
        
        # Anexos por sexo
        sexo = scraping_desempleo_sexo(path=carpeta+r"\DANE",tiempo=t)
        c_sexo = clean_sexo(path=carpeta+r"\DANE")
        
        # Anexos por región
        region = scraping_desempleo_regiones(path=carpeta+r"\DANE",tiempo=t)
        c_region = clean_regiones(path=carpeta+r"\DANE")

        # Anexos estacionalizados mensual
        estacionalizado = scraping_desempleo_estacionalizado(path=carpeta+r"\DANE",tiempo=t)
        c_estac = clean_anexos_estacionalizado(path=carpeta+r"\DANE")
            
    else:
        for i in indicadores:
            if i == 'Informalidad':
                informalidad = scraping_informalidad(path=carpeta+r"\DANE",tiempo=t)
                c_infor = clean_informalidad(path=carpeta+r"\DANE")
                
            elif i == 'Desempleo_desestacionalizado':
                
                desestacionalizado = scraping_desempleo_desetacionalizada_mensual(path=carpeta+r"\DANE",tiempo=t)
                c_deses = clean_anexos_desestacionalizado(path=carpeta+r"\DANE")
                
            elif i == 'Desempleo_por_sexo':
                sexo = scraping_desempleo_sexo(path=carpeta+r"\DANE",tiempo=t)
                c_sexo = clean_sexo(path=carpeta+r"\DANE")
                
            elif i == 'Desempleo_por_region':
                region = scraping_desempleo_regiones(path=carpeta+r"\DANE",tiempo=t)
                c_region = clean_regiones(path=carpeta+r"\DANE")
            
            elif i == 'Desempleo_estacionalizado':
                estacionalizado = scraping_desempleo_estacionalizado(path=carpeta+r"\DANE",tiempo=t)
                c_estac = clean_anexos_estacionalizado(path=carpeta+r"\DANE")
                
            else:
                print('Indicador no válido, verifique que esté escrito correctamente')

    if excel:
        excel = guardar_excel(Fuente=fuente_DANE,
                              carpeta_origen=carpeta+r"\DANE\CSV",
                              carpeta_destino=carpeta,
                              nombre_archivo='HUB_MercadoLaboral_DANE_NV',
                              hyperlinks=True)
        
    fuente_DANE.to_excel(carpeta+"\indicadores.xlsx")

    

                
    


