import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook

import os
import xlrd
import shutil
import sys

def clean_anexos_estacionalizado(path):
    
    try:
        os.mkdir(path+"\\archivos_fuente")
        os.mkdir(path+"\\CSV")
    except:
        pass
        
    ###############################################3
    # Nombre del archivo 
        
    archivos =  os.listdir(path)
    archivos.remove("archivos_fuente")
    archivos.remove("CSV")
    file_name = archivos[0]    
    
    ####################################################
    ###################################################
    
    ## MLE_TNN Mercado laboral esetacionalizado total nacional
    
    try:
        
        data = pd.read_excel(path+"/"+file_name,
                             sheet_name = "Tnal mensual")
        divisiones = ["Total Nacional"]
        df = pd.DataFrame({})

        for j in divisiones:
            l = data.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]
            df_temp = data.iloc[sup+3:sup+35,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2001-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.drop('nan',axis=1)
            df_temp1 = df_temp.iloc[:,:14].applymap(lambda x: float(x)/100)
            df_temp2 = df_temp.iloc[:,14:].applymap(lambda x: float(x)*1000)
            df_temp = pd.concat([df_temp1,df_temp2],axis=1)
            df_temp.insert(0,'División',j[:])

            df_temp.columns = [ 'División',
                                '% población en edad de trabajar ',
                                'TGP', 'TO', 'TD', 'T.D. Abierto','T.D. Oculto',
                                'Tasa de subempleo subjetivo %',
                                'Insuficiencia de horas subjetivo %',
                                'Empleo inadecuado por competencias subjetivo %',
                                'Empleo inadecuado por ingresos subjetivo %',
                                'Tasa de subempleo objetivo %',
                                'Insuficiencia de horas objetivo %',
                                'Empleo inadecuado por competencias objetivo %',
                                'Empleo inadecuado por ingresos objetivo %',
                                'Población total',
                                'Población en edad de trabajar',
                                'Población económicamente activa',
                                'Ocupados', 'Desocupados', 'Abiertos', 'Ocultos', 'Inactivos',
                                'Subempleados Subjetivos',
                                'Insuficiencia de horas subjetivo',
                                'Empleo inadecuado por competencias subjetivo',
                                'Empleo inadecuado por ingresos subjetivo',
                                'Subempleados Objetivos',
                                'Insuficiencia de horas objetivo',
                                'Empleo inadecuado por competencias objetivo',
                                'Empleo inadecuado por ingresos objetivo']

            df = df_temp

        df.to_csv(path+r"\CSV\MLE_TNN_{}_desempleo_estacionalizado_total_nacional.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        print("""
        4 Limpieza de Desempleo Eetacionalizado

        Archivo: {}
        ---------------------------
        Hojas:
        ---------------------------
        Tnal mensual
        Fecha actualizada: {}       
        """.format(file_name,fecha[-1]))
        
        
    except ValueError as ve:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(
             """
              ----
              Error
              ----

              ubicación: HUB_DAMAC\mercado_laboral
              script: limpieza_anexos_estacionalizado.py
              función: clean_anexos_estacionalizado
              linea del código: {}

              tipo: {}


              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    
    ####################################################
    ###################################################
    
    ## MLE_TCR Mercado laboral esetacionalizado total nacional, cabeceras y rural disperso
    
    try:
        
        data2 = pd.read_excel(path+"/"+file_name,
                     sheet_name = "tnal cabe ru trim movil")
        
        divisiones = ["Total Nacional","Total Cabeceras","Centros poblados y rural disperso"]
        df2 = pd.DataFrame({})

        for j in divisiones:
            l = data2.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]
            df_temp = data2.iloc[sup+3:sup+35,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2001-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.drop('nan',axis=1)
            df_temp1 = df_temp.iloc[:,:14].applymap(lambda x: float(x)/100)
            df_temp2 = df_temp.iloc[:,14:].applymap(lambda x: float(x)*1000)
            df_temp = pd.concat([df_temp1,df_temp2],axis=1)
            df_temp.insert(0,'División',j[:])

            df_temp.columns = [ 'División',
                                '% población en edad de trabajar ',
                                'TGP', 'TO', 'TD', 'T.D. Abierto','T.D. Oculto',
                                'Tasa de subempleo subjetivo %',
                                'Insuficiencia de horas subjetivo %',
                                'Empleo inadecuado por competencias subjetivo %',
                                'Empleo inadecuado por ingresos subjetivo %',
                                'Tasa de subempleo objetivo %',
                                'Insuficiencia de horas objetivo %',
                                'Empleo inadecuado por competencias objetivo %',
                                'Empleo inadecuado por ingresos objetivo %',
                                'Población total',
                                'Población en edad de trabajar',
                                'Población económicamente activa',
                                'Ocupados', 'Desocupados', 'Abiertos', 'Ocultos', 'Inactivos',
                                'Subempleados Subjetivos',
                                'Insuficiencia de horas subjetivo',
                                'Empleo inadecuado por competencias subjetivo',
                                'Empleo inadecuado por ingresos subjetivo',
                                'Subempleados Objetivos',
                                'Insuficiencia de horas objetivo',
                                'Empleo inadecuado por competencias objetivo',
                                'Empleo inadecuado por ingresos objetivo']

            df2 = pd.concat([df2,df_temp],axis=0)

        df2.to_csv(path+r"\CSV\MLE_TCR_{}_desempleo_estacionalizado_total_nacional_cabeceras_resto.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        print("""
        tnal cabe ru trim movil 
        Fecha actualizada: {}            
        """.format(fecha[-1]))
        
        
    except ValueError as ve:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(
             """
              ----
              Error
              ----

              ubicación: HUB_DAMAC\mercado_laboral
              script: limpieza_anexos_estacionalizado.py
              función: clean_anexos_estacionalizado
              linea del código: {}

              tipo: {}


              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    
    ####################################################
    ###################################################
    
    ## MLE_TCR Mercado laboral esetacionalizado areas y ciudades
    
    try:
        
        data3 = pd.read_excel(path+"/"+file_name,
                     sheet_name = "areas trim movil")
        
        divisiones = ['Total 13 ciudades y áreas metropolitanas',
                     'Bogotá',
                     'Medellín A.M.',
                     'Cali A.M.',
                     'Barranquilla A.M.',
                     'Bucaramanga A.M.',
                     'Manizales A.M.',
                     'Pasto',
                     'Pereira A.M.',
                     'Cúcuta A.M.',
                     'Ibagué',
                     'Montería',
                     'Cartagena',
                     'Villavicencio',
                     'Tunja',
                     'Florencia',
                     'Popayán',
                     'Valledupar',
                     'Quibdó',
                     'Neiva',
                     'Riohacha',
                     'Santa Marta',
                     'Armenia',
                     'Sincelejo',
                     'Total 10 ciudades',
                     'Total 23 ciudades y A.M.']
        df3 = pd.DataFrame({})

        for j in divisiones:
            l = data3.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]
            if j == 'Total 13 ciudades y áreas metropolitanas':
                df_temp = data3.iloc[sup+3:sup+35,:]
            else:
                df_temp = data3.iloc[sup+4:sup+36,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2001-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.drop('nan',axis=1)
            df_temp1 = df_temp.iloc[:,:14].applymap(lambda x: float(x)/100)
            df_temp2 = df_temp.iloc[:,14:].applymap(lambda x: float(x)*1000)
            df_temp = pd.concat([df_temp1,df_temp2],axis=1)
            df_temp.insert(0,'División',j[:])

            df_temp.columns = [ 'División',
                                '% población en edad de trabajar ',
                                'TGP', 'TO', 'TD', 'T.D. Abierto','T.D. Oculto',
                                'Tasa de subempleo subjetivo %',
                                'Insuficiencia de horas subjetivo %',
                                'Empleo inadecuado por competencias subjetivo %',
                                'Empleo inadecuado por ingresos subjetivo %',
                                'Tasa de subempleo objetivo %',
                                'Insuficiencia de horas objetivo %',
                                'Empleo inadecuado por competencias objetivo %',
                                'Empleo inadecuado por ingresos objetivo %',
                                'Población total',
                                'Población en edad de trabajar',
                                'Población económicamente activa',
                                'Ocupados', 'Desocupados', 'Abiertos', 'Ocultos', 'Inactivos',
                                'Subempleados Subjetivos',
                                'Insuficiencia de horas subjetivo',
                                'Empleo inadecuado por competencias subjetivo',
                                'Empleo inadecuado por ingresos subjetivo',
                                'Subempleados Objetivos',
                                'Insuficiencia de horas objetivo',
                                'Empleo inadecuado por competencias objetivo',
                                'Empleo inadecuado por ingresos objetivo']

            df3 = pd.concat([df3,df_temp],axis=0)

        df3.to_csv(path+r"\CSV\MLE_TAC_{}_desempleo_estacionalizado_total_areas_ciudades.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        print("""
        areas trim movil 
        Fecha actualizada: {}              
        """.format(fecha[-1]))
        
        
    except ValueError as ve:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(
             """
              ----
              Error
              ----

              ubicación: HUB_DAMAC\mercado_laboral
              script: limpieza_anexos_estacionalizado.py
              función: clean_anexos_estacionalizado
              linea del código: {}

              tipo: {}


              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    
 


    ####################################################
    ###################################################
    
    ## MLE_ORD Mercado laboral estacionalizado ocupados ramas divisiones
    
    try:
        
        data4 = pd.read_excel(path+"/"+file_name,
                         sheet_name = "ocup ramas trim tnal CIIU 4 ")
        
        divisiones = ["TOTAL NACIONAL","CABECERAS","CENTROS POBLADOS Y RURAL DISPERSO"]
        df4 = pd.DataFrame({})

        for j in divisiones:
            l = data4.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            df_temp = data4.iloc[sup+3:sup+20,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2015-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.drop('nan',axis=1)
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:])

            df_temp.columns=["Divisón",
                            "Total Ocupados",
                            "No informa",
                            "Agricultura, ganadería, caza, silvicultura y pesca",
                            "Explotación de minas y canteras",
                            "Industrias manufactureras",
                            "Suministro de electricidad gas, agua y gestión de desechos",
                            "Construcción",
                            "Comercio y reparación de vehículos",
                            "Alojamiento y servicios de comida",
                            "Transporte y almacenamiento",
                            "Información y comunicaciones",
                            "Actividades financieras y de seguros",
                            "Actividades inmobiliarias",
                            "Actividades profesionales, científicas, técnicas y servicios administrativos",
                            "Administración pública y defensa, educación y atención de la salud humana",
                            "Actividades artísticas, entretenimiento, recreación y otras actividades de servicios"
                            ]

            df4 = pd.concat([df4,df_temp],axis=0)

        df4.to_csv(path+r"\CSV\MLE_ORD_{}_desempleo_estacionalizado_ocupados_ramas_divisiones.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        print("""
        ocup ramas trim tnal CIIU 4 
        Fecha actualizada: {}              
        """.format(fecha[-1]))
        
        
    except ValueError as ve:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(
             """
              ----
              Error
              ----

              ubicación: HUB_DAMAC\mercado_laboral
              script: limpieza_anexos_estacionalizado.py
              función: clean_anexos_estacionalizado
              linea del código: {}

              tipo: {}


              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    
    
    ####################################################
    ###################################################    
    ## MLE_ Mercado laboral estacionalizado ocupados ramas ciudades
    
    try:
        
        data5 = pd.read_excel(path+"/"+file_name,
                         sheet_name = "ocu ramas trim 23 áreas CIIU 4")
        
        divisiones = ['OCUPADOS 13 CIUDADES Y ÁREAS METROPOLITANAS',
                     'BOGOTÁ',
                     'MEDELLÍN A.M.',
                     'CALI  A.M.',
                     'BARRANQUILLA A.M.',
                     'BUCARAMANGA A.M.',
                     'MANIZALES A.M.',
                     'PASTO',
                     'PEREIRA A.M.',
                     'CÚCUTA A.M.',
                     'IBAGUÉ',
                     'MONTERÍA',
                     'CARTAGENA',
                     'VILLAVICENCIO',
                     'TUNJA',
                     'FLORENCIA',
                     'POPAYÁN',
                     'VALLEDUPAR',
                     'QUIBDÓ',
                     'NEIVA',
                     'RIOHACHA',
                     'SANTA MARTA',
                     'ARMENIA',
                     'SINCELEJO']
        df5 = pd.DataFrame({})

        for j in divisiones:
            l = data5.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            df_temp = data5.iloc[sup+4:sup+21,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2015-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.drop('nan',axis=1)
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:])

            df_temp.columns=["Divisón",
                            "Total Ocupados",
                            "No informa",
                            "Agricultura, ganadería, caza, silvicultura y pesca",
                            "Explotación de minas y canteras",
                            "Industrias manufactureras",
                            "Suministro de electricidad gas, agua y gestión de desechos",
                            "Construcción",
                            "Comercio y reparación de vehículos",
                            "Alojamiento y servicios de comida",
                            "Transporte y almacenamiento",
                            "Información y comunicaciones",
                            "Actividades financieras y de seguros",
                            "Actividades inmobiliarias",
                            "Actividades profesionales, científicas, técnicas y servicios administrativos",
                            "Administración pública y defensa, educación y atención de la salud humana",
                            "Actividades artísticas, entretenimiento, recreación y otras actividades de servicios"
                            ]

            df5 = pd.concat([df5,df_temp],axis=0)

        df5.to_csv(path+r"\CSV\MLE_ORC_{}_desempleo_estacionalizado_ocupados_ramas_ciudades.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        print("""
        ocu ramas trim 23 áreas CIIU 4
        Fecha actualizada: {}             
        """.format(fecha[-1]))
        
        
    except ValueError as ve:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(
             """
              ----
              Error
              ----

              ubicación: HUB_DAMAC\mercado_laboral
              script: limpieza_anexos_estacionalizado.py
              función: clean_anexos_estacionalizado
              linea del código: {}

              tipo: {}


              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    
    ####################################################
    ###################################################
    
    ## MLE_OPD Mercado laboral estacionalizado ocupados posición ocupacional divisiones
    
    try:
        
        data6 = pd.read_excel(path+"/"+file_name,
                         sheet_name = "ocup posc trim tnal ")
        
        divisiones = ["TOTAL NACIONAL","CABECERAS","CENTROS POBLADOS Y RURAL DISPERSO"]
        df6 = pd.DataFrame({})

        for j in divisiones:
            l = data6.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            df_temp = data6.iloc[sup+4:sup+15,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2001-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.drop('nan',axis=1)
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:])

            df_temp.columns=["División",
                            "Total Ocupados",
                            "Obrero, empleado particular",  
                            "Obrero, empleado del gobierno", 
                            "Empleado doméstico", 
                            "Trabajador por cuenta propia", 
                            "Patrón o empleador",
                            "Trabajador familiar sin remuneración", 
                            "Trabajador sin remuneración en empresas de otros hogares",
                            "Jornalero o Peón",
                            "Otro"
                            ]

            df6 = pd.concat([df6,df_temp],axis=0)

        df6.to_csv(path+r"\CSV\MLE_OPD_{}_desempleo_estacionalizado_ocupados_posicion_ocupacional_divisiones.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        print("""
        ocup posc trim tnal 
        Fecha actualizada: {}             
        """.format(fecha[-1]))
        
        
        
    except ValueError as ve:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(
             """
              ----
              Error
              ----

              ubicación: HUB_DAMAC\mercado_laboral
              script: limpieza_anexos_estacionalizado.py
              función: clean_anexos_estacionalizado
              linea del código: {}

              tipo: {}


              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
  

    ####################################################
    ###################################################
    
    ## MLE_OPC Mercado laboral estacionalizado ocupados posición ocupacional ciudades
    
    try:
        
        data7 = pd.read_excel(path+"/"+file_name,
                         sheet_name = "ocu posc trim 23 áreas")
        
        divisiones = ['TOTAL 13 CIUDADES Y ÁREAS METROPOLITANAS',
                     'BOGOTÁ',
                     'MEDELLÍN A.M.',
                     'CALI A.M.',
                     'BARRANQUILLA A.M.',
                     'BUCARAMANGA A.M.',
                     'MANIZALES A.M.',
                     'PASTO',
                     'PEREIRA A.M.',
                     'CÚCUTA A.M.',
                     'IBAGUÉ',
                     'MONTERÍA',
                     'CARTAGENA',
                     'VILLAVICENCIO',
                     'TUNJA',
                     'FLORENCIA',
                     'POPAYÁN',
                     'VALLEDUPAR',
                     'QUIBDÓ',
                     'NEIVA',
                     'RIOHACHA',
                     'SANTA MARTA',
                     'ARMENIA',
                     'SINCELEJO']
        df7 = pd.DataFrame({})

        for j in divisiones:
            l = data7.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            df_temp = data7.iloc[sup+3:sup+14,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2001-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.drop('nan',axis=1)
            df_temp = df_temp.applymap(lambda x: str(x).replace('enan',''))
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:])
            
            df_temp.columns=["División",
                            "Total Ocupados",
                            "Obrero, empleado particular",  
                            "Obrero, empleado del gobierno", 
                            "Empleado doméstico", 
                            "Trabajador por cuenta propia", 
                            "Patrón o empleador",
                            "Trabajador familiar sin remuneración", 
                            "Trabajador sin remuneración en empresas de otros hogares",
                            "Jornalero o Peón",
                            "Otro"
                            ]

            df7 = pd.concat([df7,df_temp],axis=0)

        df7.to_csv(path+r"\CSV\MLE_OPC_{}_desempleo_estacionalizado_ocupados_posicion_ocupacional_ciudades.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        print("""
        ocu posc trim 23 áreas 
        Fecha actualizada: {}       
        ----------------------------------------------
        Guardado en: {}         
        """.format(fecha[-1],path+r"\CSV"))
        
        
    except ValueError as ve:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(
             """
              ----
              Error
              ----

              ubicación: HUB_DAMAC\mercado_laboral
              script: limpieza_anexos_estacionalizado.py
              función: clean_anexos_estacionalizado
              linea del código: {}

              tipo: {}


              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    
    finally:
        shutil.move(path+r"\{}".format(file_name),path+r"\archivos_fuente\{}".format(file_name))