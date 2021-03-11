from HUB import HUB_DAMAC

laboral = HUB_DAMAC.mercado_laboral()
laboral.actualizar(carpeta=r"D:\Desktop\Laboral",actualizar_todo=True, excel=True,hipervinculos=True)
# print(laboral.descripcion)
# laboral.fuente_laboral
# guardar_excel(Fuente=laboral.fuente_laboral,carpeta_origen=r"D:\Desktop\Laboral",carpeta_destino=r"D:\Desktop\Laboral", nombre_archivo='hub_laboral',hyperlinks=True)
#laboral.actualizar(carpeta=r"D:\Desktop\Laboral",actualizar_todo=False,indicadores=['Tasa de desempleo','Tasa de ocupación'],excel=True,hipervinculos=True)
# laboral.actualizar(carpeta=r"D:\Desktop\Laboral",actualizar_todo=False,indicadores=['Tasa de desempleo','Informalidad'],excel=True,hipervinculos=True)


import os
import pandas as pd 

path = "D:/Desktop/HUB/mercado_laboral/"
archivos =  pd.Series(os.listdir(path))
dane_informalidad_nombre = archivos[archivos.str.contains('informalidad')][0]
for i in archivos:
    if i == dane_informalidad_nombre:
        # print(i)
        data= pd.read_excel(path+'\{}'.format(i),sheet_name=3)

        # index = data[data.iloc[:,0].str.contains('23 Ciudades').fillna(False)].index[0]

        # data_23_ciudades = data.iloc[index:,:]

        # periodo = data_23_ciudades[data_23_ciudades.iloc[:,1].str.contains('Ene-').fillna(False)].dropna(axis=1).T
        # periodo.columns = ['Periodo']
        # periodo.reset_index(level=0,drop=True,inplace = True)

        # fecha = pd.date_range(start='2007',freq='M',periods=len(periodo))
        # periodo['Fecha'] = fecha

        # tasa = data_23_ciudades[data_23_ciudades.iloc[:,0].str.contains('Ocupados').fillna(False)].dropna(axis=1).T.iloc[1:,:]

        # tasa.columns = ['Ocupados Informales']
        # tasa['Ocupados Informales'] = tasa['Ocupados Informales'].astype('float')
        # tasa.reset_index(level=0,drop=True,inplace=True)

        # tasa_informalidad = pd.concat([periodo,tasa],axis=1)
        # print('Limpieza Exitosa')
data