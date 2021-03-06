import pandas as pd
import numpy as np
import os

def clean_informalidad(path):
    try:
        archivos =  pd.Series(os.listdir(path))
        dane_informalidad_nombre = archivos[archivos.str.contains('informalidad')][0]
        for i in archivos:
            if i == dane_informalidad_nombre:
                try:  
                    data= pd.read_excel(path+'\{}'.format(i),sheet_name=3)

                    index = data[data.iloc[:,0].str.contains('23 Ciudades').fillna(False)].index[0]

                    data_23_ciudades = data.iloc[index:,:]

                    periodo = data_23_ciudades[data_23_ciudades.iloc[:,1].str.contains('Ene-').fillna(False)].dropna(axis=1).T
                    periodo.columns = ['Periodo']
                    periodo.reset_index(level=0,drop=True,inplace = True)

                    fecha = pd.date_range(start='2007',freq='M',periods=len(periodo))
                    periodo['Fecha'] = fecha

                    tasa = data_23_ciudades[data_23_ciudades.iloc[:,0].str.contains('Ocupados').fillna(False)].dropna(axis=1).T.iloc[1:,:]

                    tasa.columns = ['Ocupados Informales']
                    tasa['Ocupados Informales'] = tasa['Ocupados Informales'].astype('float')
                    tasa.reset_index(level=0,drop=True,inplace=True)

                    tasa_informalidad = pd.concat([periodo,tasa],axis=1)
                    print('Limpieza Exitosa')
                except:
                    print("""No se pudo limpiar correctamente la tasa de informalidad, 
                    asegurese de instalar xlrd use pip install xlrd y luego conda install xlrd en la consola""")

        tasa_informalidad.to_csv(path+'\Informalidad.csv',sep=';',decimal=',')  
        os.remove(path+'\{}'.format(dane_informalidad_nombre))
    except:
        pass





