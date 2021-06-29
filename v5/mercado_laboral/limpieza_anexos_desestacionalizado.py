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

def clean_anexos_desestacionalizado(path):
    
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
    
    ## MLD_TNN Mercado laboral desetacionalizado total nacional
    
    try:
        data = pd.read_excel(path+"/"+file_name,
                     sheet_name = "Tnal mensual")
        divisiones = ["Total Nacional desestacionalizada"]
        df = pd.DataFrame({})

        for j in divisiones:
            l = data.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]
            df_temp = data.iloc[sup+3:sup+11,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2001-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.drop('nan',axis=1)
            df_temp1 = df_temp.iloc[:,:3].applymap(lambda x: float(x)/100)
            df_temp2 = df_temp.iloc[:,3:].applymap(lambda x: float(x)*1000)
            df_temp = pd.concat([df_temp1,df_temp2],axis=1)
            df_temp.insert(0,'División',j[:])

            df = pd.concat([df,df_temp],axis=0)

        df.to_csv(path+r"\CSV\MLD_TNN_{}_desempleo_desestacionalizado_total_nacional.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        print("""
        4 Limpieza de Desempleo desetacionalizado

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
              script: limpieza_anexos_desestacionalizado.py
              función: clean_anexos_desestacionalizado
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
    
    ## MLD_TCR Mercado laboral desetacionalizado total nacional, cabeceras y resto
    
    try:
        data2 = pd.read_excel(path+"/"+file_name,
                     sheet_name = "tnal cabe ru trim movil")
        
        divisiones = ["Total Nacional desestacionalizada","Total Cabeceras desestacionalizada",
                      "Centros poblados y rural disperso desestacionalizada"]
        df2 = pd.DataFrame({})

        for j in divisiones:
            l = data2.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]
            df_temp = data2.iloc[sup+3:sup+11,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2001-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.drop('nan',axis=1)
            df_temp1 = df_temp.iloc[:,:3].applymap(lambda x: float(x)/100)
            df_temp2 = df_temp.iloc[:,3:].applymap(lambda x: float(x)*1000)
            df_temp = pd.concat([df_temp1,df_temp2],axis=1)
            df_temp.insert(0,'División',j[:])

            df2 = pd.concat([df2,df_temp],axis=0)

        df2.to_csv(path+r"\CSV\MLD_TCR_{}_desempleo_desestacionalizado_total_nacional_cabeceras_resto.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        print("""
        tnal cabe ru trim movil 
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
              script: limpieza_anexos_desestacionalizado.py
              función: clean_anexos_desestacionalizado
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