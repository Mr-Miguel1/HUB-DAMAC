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

def clean_informalidad(path):
    
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
    
    ## MLI-TNN TOTAL NACIONAL

    
    try:

        # Tener en cuenta el nombre de la hoja en el archivo excel del dane
        data = pd.read_excel(path+"/"+file_name,
                            sheet_name = 'Prop informalidad')

        # Indices de las filas de cada ciudad y A.M
        index_ciudad = np.arange(9,154,6)
        
        # operamos con iloc (segurarse de que los indices sean correctos)
        df = data.iloc[index_ciudad+4,1:].T
        df = df.applymap(lambda x: float(x)/100)
        df.columns = data.iloc[index_ciudad,0].tolist()

        df.insert(0,'Año',data.iloc[10,1:])
        df = df.fillna(method='ffill')

        df.insert(1,'Trimestre Móvil',data.iloc[11,1:])

        fecha = pd.date_range(start='2007-01-01',periods=df.shape[0],freq='M',name='Fecha')
        df.set_index(fecha,drop=True,inplace=True)

        df.to_csv(path+r"\CSV\MLI_TNN_{}_tasa_total_nacional.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        print("""
        
        1 Limpieza de Informalidad exitosa
        
        Archivo: {}
        ---------------------------
        Hojas:
        ---------------------------
        Prop informalidad
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
              script: limpieza_informalidad.py
              función: clean_informalidad
              linea del código: {}
              
              tipo: {}
              
              
              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    

    ####################################################################
    ####################################################################
    
    ## MLI-CIU  CIUDADES
    
    
    try:
        # Tener en cuenta el nombre de la hoja en el archivo excel del dane        
        data2 =  pd.read_excel(path+"/"+file_name,
                     sheet_name = 'Ciudades')
        data2 = data2.applymap(lambda x: str(x).upper())
        data2 = data2.applymap(lambda x: str(x).replace(' ','_'))

        index_ciudades = []
        index_ocupados = []
        index_formales = []
        index_informales = []
        años = [i for i in data2.iloc[9,1:]]
        trimestres = [i for i in data2.iloc[10,1:]]
        for ix,i in enumerate(data2.iloc[:,0]):
            if i == 'OCUPADOS':
                index_ciudades.append(ix-2)
                index_ocupados.append(ix)
                index_formales.append(ix+1)
                index_informales.append(ix+2)


        fecha = pd.date_range(start='2007-01-01',periods=len(data2.iloc[0,1:]),freq='M')
        ciudades = data2.iloc[index_ciudades,0].tolist()

        dic = {}
        df2 = pd.DataFrame({})
        for ciudad in ciudades:
            for date in fecha:
                dic['Ciudad'] = ciudad
                dic['Fecha'] = fecha

            df_temp = pd.DataFrame(dic,index=np.arange(0,len(fecha),1))   
            df2 = pd.concat([df2,df_temp],axis=0)

        ocupados = np.array(data2.iloc[index_ocupados,1:]).reshape(-1,1)
        formales = np.array(data2.iloc[index_formales,1:]).reshape(-1,1)
        informales = np.array(data2.iloc[index_informales,1:]).reshape(-1,1)

        df2.insert(0,'Trimestre Móvil',trimestres*25)
        df2.set_index('Fecha',drop=True,inplace=True)
        df2.applymap(lambda x: x.replace('_',''))

        df2['Ocupados'] = [float(i)*1000 for i in ocupados]
        df2['Formales'] = [float(i)*1000 for i in formales]
        df2['Informales'] = [float(i)*1000 for i in informales]
        
        
        df2.to_csv(path+r"\CSV\MLI_CIU_{}_ciudades.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        print(""" 
        Ciudades
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
              script: limpieza_informalidad.py
              función: clean_informalidad
              linea del código: {}
              
              tipo: {}
              
              
              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    
    ####################################################################
    ####################################################################
    
    ## MLI-SEX  SEXO
        
    try:
        data3 =  pd.read_excel(path+"/"+file_name,
                       sheet_name = 'Sexo')
        divisiones = ["Ocupados 13 ciudades y áreas metropolitanas","Ocupados 23 ciudades y áreas metropolitanas"]
        df3 = pd.DataFrame({})

        for j in divisiones:
            l = data3.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            df_temp = data3.iloc[sup:sup+9,:]
            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2007-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:20]+' A.M')

            df_temp.columns = ['División',
                               'Total',
                               'Total Informales',
                               'Total Formales',
                               "Total Hombres",
                               "Hombres Informales",
                               "Hombres Formales",
                               "Total Mujeres",
                               "Mujeres Informales",
                               "Mujeres Formales"]
            df3 = pd.concat([df3,df_temp],axis=0)
            
        df3.to_csv(path+r"\CSV\MLI_SEX_{}_sexo.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        print(""" 
        Sexo
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
              script: limpieza_informalidad.py
              función: clean_informalidad
              linea del código: {}
              
              tipo: {}
              
              
              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    
    ####################################################################
    ####################################################################
    
    ## MLI-EDU  EDUCACION
    
    try:
        
        data4 = pd.read_excel(path+"/"+file_name,
                     sheet_name = 'Educación ')
        divisiones = ["Ocupados 13 ciudades y áreas metropolitanas","Ocupados 23 ciudades y áreas metropolitanas"]
        df4 = pd.DataFrame({})

        for j in divisiones:
            l = data4.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            df_temp = data4.iloc[sup:sup+18,:]
            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2007-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:20]+' A.M')

            df_temp.columns = ['División',
                               'Total',
                               'Total Ninguno',
                               'Total Primaria',
                               "Total Secundaria",
                               "Total Superior",
                               "Total No Informa",
                               'Informales',
                               'Informales Ninguno',
                               'Informales Primaria',
                               "Informales Secundaria",
                               "Informales Superior",
                               "Informales No Informa",
                               'Formales',
                               'Formales Ninguno',
                               'Formales Primaria',
                               "Formales Secundaria",
                               "Formales Superior",
                               "Formales No Informa",
                              ]
            df4 = pd.concat([df4,df_temp],axis=0)

        df4.to_csv(path+r"\CSV\MLI_EDU_{}_educacion.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        print(""" 
        Educación
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
              script: limpieza_informalidad.py
              función: clean_informalidad
              linea del código: {}
              
              tipo: {}
              
              
              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    
    ####################################################################
    ####################################################################
    
    ## MLI-PSC  POSICIÓN OCUPACIONAL
    
    try:

        data5 = pd.read_excel(path+"/"+file_name,
                             sheet_name = 'posi ocupacional')
        
        divisiones = ["Ocupados 13 ciudades y áreas metropolitanas","Ocupados 23 ciudades y áreas metropolitanas"]
        df5 = pd.DataFrame({})

        for j in divisiones:
            l = data5.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            df_temp = data5.iloc[sup:sup+29,:]
            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2007-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:20]+' A.M')

            df_temp.columns = ['División',
                               'Total',
                               'Total Emp. particular',
                               'Total Emp. gobierno',
                               "Total Emp. domestico",
                               "Total Cuenta propia",
                               "Total Patron o empleador",
                               "Total Trabajador familiar sin remuneración",
                               "Total Trabajador sin remuneración en empresas de otros hogares",
                               "Total Jornalero o Peón",
                               "Total Otro",
                               'Informales',
                               'Informales Emp. particular',
                               "Informales Emp. domestico",
                               "Informales Cuenta propia",
                               "Informales Patron o empleador",
                               "Informales Trabajador familiar sin remuneración",
                               "Informales Trabajador sin remuneración en empresas de otros hogares",
                               "Informales Jornalero o Peón",
                               "Informales Otro",
                               'Formales',
                               'Formales Emp. particular',
                               'Formales Emp. gobierno',
                               "Formales Emp. domestico",
                               "Formales Cuenta propia",
                               "Formales Patron o empleador",
                               "Formales Trabajador familiar sin remuneración",
                               "Formales Trabajador sin remuneración en empresas de otros hogares",
                               "Formales Jornalero o Peón",
                               "Formales Otro"]
            df5 = pd.concat([df5,df_temp],axis=0)

        df5.to_csv(path+r"\CSV\MLI_PSC_{}_posicion_ocupacional.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        print(""" 
        posi ocupacional
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
              script: limpieza_informalidad.py
              función: clean_informalidad
              linea del código: {}
              
              tipo: {}
              
              
              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
 
    ####################################################################
    ####################################################################
    
    ## MLI-CII Ramas CIIU 4
    
    
    try:
        data6 = pd.read_excel(path+"/"+file_name,
                              sheet_name = 'ramas actividad CIIU 4 A.C')
        
        divisiones = ["Total 13 áreas","Ocupados 23 ciudades y áreas metropolitanas"]
        df6 = pd.DataFrame({})

        for j in divisiones:
            l = data6.iloc[:,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            df_temp = data6.iloc[sup:sup+32,:]
            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2015-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:20]+' A.M')

            df_temp.columns =['División',
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
                             'Actividades artísticas, entretenimiento, recreación y otras actividades de servicios',
                             'Informales',
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

            df6 = pd.concat([df6,df_temp],axis=0)

        df6.to_csv(path+r"\CSV\MLI_CII_{}_ramas_ciiu4.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        print(""" 
        ramas ciiu 4
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
              script: limpieza_informalidad.py
              función: clean_informalidad
              linea del código: {}
              
              tipo: {}
              
              
              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    
    
    ####################################################################
    ####################################################################
    
    ## MLI-CII Seguridad Social Cantidad de personas
    
    try:
        
        data7 = pd.read_excel(path+"/"+file_name,
                     sheet_name = 'Seguridad social T. nal')
        
        divisiones = ["Total Nacional","Cabeceras","Centro poblado y rural disperso"]
        df7 = pd.DataFrame({})

        for j in divisiones:
            l = data7.iloc[:49,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            df_temp = data7.iloc[sup+3:sup+13,:]
            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2007-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.applymap(lambda x: float(x)*1000)
            df_temp.insert(0,'División',j[:20])


            df7 = pd.concat([df7,df_temp],axis=0)

        df7.to_csv(path+r"\CSV\MLI_SSC_{}_seguridad_social_cantidad_personas.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        print(""" 
        Seguridad social T. nal cantidad personas
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
              script: limpieza_informalidad.py
              función: clean_informalidad
              linea del código: {}
              
              tipo: {}
              
              
              -----
              Información del error:
              -----
              + {} 
              """.format(exc_tb.tb_lineno,exc_type,ve))
    else:
        pass
    
    
    ####################################################################
    ####################################################################
    
    ## MLI-CII Seguridad Social Porcentaje de personas    
    try:
        
        data8 = pd.read_excel(path+"/"+file_name,
                     sheet_name = 'Seguridad social T. nal')
        
        divisiones = ["Total nacional","Cabeceras","Centros poblados y rural disperso"]
        df8 = pd.DataFrame({})

        for j in divisiones:
            l = data8.iloc[50:94,0]
            sup = l[l.str.contains(j).fillna(False)].index[0]

            df_temp = data8.iloc[sup+3:sup+13,:]
            df_temp = df_temp.T.dropna(how='all',axis=0)
            df_temp.columns = df_temp.iloc[0,:]

            df_temp.reset_index(level=0,drop=True,inplace=True)
            df_temp = df_temp.drop([0],axis=0)

            fecha = pd.date_range(start='2007-01-01',periods=len(df_temp),freq='M',name='Fecha')

            df_temp.set_index(fecha,drop=True,inplace=True)
            df_temp = df_temp.applymap(lambda x: float(x)/100)
            df_temp.insert(0,'División',j[:20])


            df8 = pd.concat([df8,df_temp],axis=0)

        df8.to_csv(path+r"\CSV\MLI_SSP_{}_seguridad_social_porcentaje_personas.csv".format(file_name[:-5]),
                 encoding = 'utf-8',
                 decimal = ',',
                 sep = ';')
        
        
        print(""" 
        Seguridad social T. nal porcentaje personas
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
              script: limpieza_informalidad.py
              función: clean_informalidad
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