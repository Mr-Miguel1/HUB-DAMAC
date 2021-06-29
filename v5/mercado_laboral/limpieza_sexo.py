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

def clean_sexo(path):
    
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
    
    ## MLS-TSN Mercado Laboral desempleo por sexo total nacional
    
    try:
        data = pd.read_excel(path+"/"+file_name,
                             sheet_name = 'P y T N')
        divisiones = ["Total Nacional","HOMBRES","MUJERES"]
        df = pd.DataFrame({})

        for j in divisiones:
            l = data.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            if j == "Total Nacional":
                 df_temp = data.iloc[sup+5:sup+32,:]
            else:
                df_temp = data.iloc[sup+3:sup+30,:]

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


            df = pd.concat([df,df_temp],axis=0)

        df.to_csv(path+r"\CSV\MLS_TSN_{}_desempleo_sexo_total_nacional.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        
        
        print("""
        3 Limpieza de Desempleo por Sexo

        Archivo: {}
        ---------------------------
        Hojas:
        ---------------------------
        P y T N 
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
              script: limpieza_sexo.py
              función: clean_sexo
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
    ## MLS-TSC Mercado Laboral desempleo por sexo total cabeceras
    
    try:
        data2 = pd.read_excel(path+"/"+file_name,
                             sheet_name = 'P y T Cab')
        divisiones = ["Total Cabeceras","HOMBRES","MUJERES"]
        df2 = pd.DataFrame({})

        for j in divisiones:
            l = data2.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            if j == "Total Cabeceras":
                 df_temp = data2.iloc[sup+5:sup+32,:]
            else:
                df_temp = data2.iloc[sup+3:sup+30,:]

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


            df2 = pd.concat([df2,df_temp],axis=0)

        df2.to_csv(path+r"\CSV\MLS_TSC_{}_desempleo_sexo_total_cabeceras.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        
        
        print(""" 
        P y T Cab
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
              script: limpieza_sexo.py
              función: clean_sexo
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
    ## MLS-TSC Mercado Laboral desempleo por sexo total resto
    
    try:
        data3 = pd.read_excel(path+"/"+file_name,
                             sheet_name = 'P y T Resto')
        divisiones = ["Total Centros poblados y rural disperso","HOMBRES","MUJERES"]
        df3 = pd.DataFrame({})

        for j in divisiones:
            l = data3.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            if j == "Total Centros poblados y rural disperso":
                 df_temp = data3.iloc[sup+5:sup+32,:]
            else:
                df_temp = data3.iloc[sup+3:sup+30,:]

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


            df3 = pd.concat([df3,df_temp],axis=0)

        df3.to_csv(path+r"\CSV\MLS_TSR_{}_desempleo_sexo_total_resto.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        
        
        print(""" 
        P y T Resto
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
              script: limpieza_sexo.py
              función: clean_sexo
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
    ## MLS-OSO Mercado Laboral ocupados por sexo posicion ocupacional
    
    try:
        data4 = pd.read_excel(path+"/"+file_name,
                     sheet_name = 'Pos ocup N')
        
        divisiones = ["Total Nacional","HOMBRES","MUJERES"]
        df4 = pd.DataFrame({})

        for j in divisiones:
            l = data4.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            if j == "Total Nacional":
                 df_temp = data4.iloc[sup+7:sup+16,:]
            else:
                df_temp = data4.iloc[sup+4:sup+13,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2001-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:])

            df_temp.columns = ['División',
                                 'Total',
                                 'Obrero, empleado particular ',
                                 'Obrero, empleado del gobierno',
                                 'Empleado doméstico ',
                                 'Trabajador por cuenta propia',
                                 'Patrón o empleador',
                                 'Trabajador familiar sin remuneración*',
                                 'Jornalero o peón',
                                 'Otro']

            df4 = pd.concat([df4,df_temp],axis=0)

        df4.to_csv(path+r"\CSV\MLS_OSO_{}_ocupados_sexo_posicion_ocupacional.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        print(""" 
        Pos ocup N
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
              script: limpieza_sexo.py
              función: clean_sexo
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
    ## MLS-OSR Mercado Laboral ocupados por sexo ramas ciiu 4 
              
    try:
        data5 = pd.read_excel(path+"/"+file_name,
                             sheet_name = 'Ramas CIIU 4 N')
        divisiones = ["Total Nacional ","HOMBRES","MUJERES"]
        df5 = pd.DataFrame({})

        for j in divisiones:
            l = data5.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            if j == "Total Nacional ":
                 df_temp = data5.iloc[sup+8:sup+24,:]
            else:
                df_temp = data5.iloc[sup+4:sup+20,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2015-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:])
            df_temp.columns = [ 'División', 
                                 'Total',
                                 'No informa',
                                 'Agricultura, ganadería, caza, silvicultura y pesca',
                                 'Explotación de minas y canteras',
                                 'Industrias manufactureras',
                                 'Suministro de electricidad gas, agua y gestión de desechos',
                                 'Construcción',
                                 'Comercio y reparación de vehículos',
                                 'Alojamiento y servicios de comida',
                                 'Transporte y almacenamiento',
                                 'Información y comunicaciones',
                                 'Actividades financieras y de seguros',
                                 'Actividades inmobiliarias',
                                 'Actividades profesionales, científicas, técnicas y servicios administrativos',
                                 'Administración pública y defensa, educación y atención de la salud humana',
                                 'Actividades artísticas, entretenimiento, recreación y otras actividades de servicios']

            df5 = pd.concat([df5,df_temp],axis=0)

        df5.to_csv(path+r"\CSV\MLS_OSR_{}_ocupados_sexo_ramas_ciiu4.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        print(""" 
        Ramas CIIU 4 N
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
              script: limpieza_sexo.py
              función: clean_sexo
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
    ## MLS-ISS Mercado Laboral inactivos por sexo
    
    try:
        data6 = pd.read_excel(path+"/"+file_name,
                             sheet_name = 'Inact N')
        divisiones = ["Total Nacional ","HOMBRES","MUJERES"]
        df6 = pd.DataFrame({})

        for j in divisiones:
            l = data6.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            if j == "Total Nacional ":
                 df_temp = data6.iloc[sup+7:sup+11,:]
            else:
                df_temp = data6.iloc[sup+4:sup+8,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2001-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:])
            df_temp.columns = ['División',
                               'Total',
                               'Estudiando ',
                               'Oficios del Hogar ',
                               'Otros']

            df6 = pd.concat([df6,df_temp],axis=0)

        df6.to_csv(path+r"\CSV\MLS_ISS_{}_inactivos_sexo.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        print(""" 
        Inact N
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
              script: limpieza_sexo.py
              función: clean_sexo
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
    ## MLS-MSD Mercado Laboral mensual desetacionalizado por sexo
    
    try:
        data7 = pd.read_excel(path+"/"+file_name,
                             sheet_name = 'Mensual Desest. Sexo')
        
        divisiones = ["Hombres","Mujeres"]
        df7 = pd.DataFrame({})

        for j in divisiones:
            l = data7.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]
            df_temp = data7.iloc[sup+3:sup+10,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2007-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.drop('nan',axis=1)
            df_temp1 = df_temp.iloc[:,:3].applymap(lambda x: float(x)/100)
            df_temp2 = df_temp.iloc[:,3:].applymap(lambda x: float(x)*1000)
            df_temp = pd.concat([df_temp1,df_temp2],axis=1)
            df_temp.insert(0,'División',j[:])

            df7 = pd.concat([df7,df_temp],axis=0)

        df7.to_csv(path+r"\CSV\MLS_MSD_{}_desempleo_sexo_mensual_desestacionalizado.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        print(""" 
        Mensual Desest. Sexo
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
              script: limpieza_sexo.py
              función: clean_sexo
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
    ## MLS-TSD Mercado Laboral trimestre movil desetacionalizado por sexo
    
    try:
        data8 = pd.read_excel(path+"/"+file_name,
                             sheet_name = "Trimestre movil Desest. Sexo")
        
        divisiones = ["Hombres","Mujeres"]
        df8 = pd.DataFrame({})

        for j in divisiones:
            l = data8.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]
            df_temp = data8.iloc[sup+3:sup+12,:]

            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]
            df_temp.columns = [str(i) for i in df_temp.columns]
            df_temp = df_temp.applymap(lambda x: str(x).replace('-','nan'))

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2007-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.drop('nan',axis=1)
            df_temp1 = df_temp.iloc[:,1:4].applymap(lambda x: float(x)/100)
            df_temp2 = df_temp.iloc[:,4:].applymap(lambda x: float(x)*1000)
            df_temp = pd.concat([df_temp1,df_temp2],axis=1)
            df_temp.insert(0,'División',j[:])

            df8 = pd.concat([df8,df_temp],axis=0)

        df8.to_csv(path+r"\CSV\MLS_TSD_{}_desempleo_sexo_trimestremovil_desestacionalizado.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        print(""" 
        Trimestre movil Desest. Sexo
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
              script: limpieza_sexo.py
              función: clean_sexo
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