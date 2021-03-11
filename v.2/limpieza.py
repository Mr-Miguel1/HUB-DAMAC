import pandas as pd
import numpy as np
import os
import xlrd

class limpieza_mercado_laboral():
    
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

    def clean_desempleo_empleo_mensual(path):
        try:
            archivos = pd.Series(os.listdir(path))
            dane_des_emp_mensual_nombre = archivos[archivos.str.contains('anexo_desestacionalizado_empleo')][0]

            for i in archivos:
                if i ==  dane_des_emp_mensual_nombre:
                    try:
                        data = load_workbook(path+"\{}".format(i))
                        sheets = pd.Series(data.sheetnames).str.lower()
                        tnal_mensual_index = sheets[sheets.str.contains('tnal mensual')].index[0]
                    except:
                        data = xlrd.open_workbook_xls(path+"\{}".format(i))
                        sheets = pd.Series(data.sheet_names()).str.lower()
                        tnal_mensual_index = sheets[sheets.str.contains('tnal mensual')].index[0]

                    df = pd.read_excel(path+"\{}".format(i),sheet_name=tnal_mensual_index)

                    ind = ['tgp','to','td','ocupados','desocupados','inactivos']
                    series = pd.DataFrame({})
                    for i in ind:
                        ser_index = df[df.applymap(lambda x: str(x).lower() == i)].dropna(how='all',axis=0).index[0]
                        ser = df.iloc[ser_index,1:].rename(i)
                        if i == 'ocupados' or i == 'desocupados' or i == 'inactivos':
                            ser = ser.set_axis(pd.date_range(start='2001-01-01',periods=len(ser),freq='M')).astype('float')*1000      
                        else:
                            ser = ser.set_axis(pd.date_range(start='2001-01-01',periods=len(ser),freq='M')).astype('float')

                        series[i] = ser

            series.to_csv(path+"\desempleo_desest_mensual.csv",sep=';',decimal=',')
            os.remove(path+'\{}'.format(dane_des_emp_mensual_nombre))

        except:
            print('El : {} no se pudo limpiar correctamente'.format(dane_des_emp_mensual_nombre))
            pass


    def clean_desempleo_empleo_sexo(path):
    archivos = pd.Series(os.listdir(path))
    dane_sexo_nombre = archivos[archivos.str.contains('anexo_sexo_')][0]

    for i in archivos:
        if i == dane_sexo_nombre:
            try: 
                data = load_workbook(path+"\{}".format(i))
                sheets = data.sheetnames
            except:
                data = xlrd.open_workbook_xls(path+"\{}".format(i))
                sheets = data.sheet_names()

            sheets = pd.Series(sheets).str.lower().str.replace(' ','')
            tnal_nacional_sexo_index = sheets[sheets.str.contains('pytn')].index[0]

            df = pd.read_excel(path+"\{}".format(i),sheet_name=tnal_nacional_sexo_index)


            ind = [i.lower().replace(' ','_') for i in pd.Series("""% población en edad de trabajar 
TGP
TO
TD
T.D. Abierto
T.D. Oculto
Población total
Población en edad de trabajar
Población económicamente activa
Ocupados
Desocupados
Abiertos
Ocultos
Inactivos""").str.split('\n')[0]]


            series_tnac = pd.DataFrame({})
            series_hombres = pd.DataFrame({})
            series_mujeres = pd.DataFrame({})

            for i in ind:

                # Total Nacional
                ser_index_nac = df[df.applymap(lambda x: str(x).lower().replace(' ','_') == i)].dropna(how='all',axis=0).index[0]

                if ser_index_nac:
                    ser = df.iloc[ser_index_nac,1:].rename(i) 
                    if i == 'tgp' or i== 'to' or i == 'td' or i == 't.d._abierto' or i == 't.d._oculto':
                        ser = ser.set_axis(pd.date_range(start='2001-01-01',periods=len(ser),freq='M')).astype('float')
                    else:
                        ser = ser.set_axis(pd.date_range(start='2001-01-01',periods=len(ser),freq='M')).astype('float')*1000
                    series_tnac[i] = ser

                #Hombres
                ser_index_hom = df[df.applymap(lambda x: str(x).lower().replace(' ','_') == i)].dropna(how='all',axis=0).index[1]

                if ser_index_hom:
                    ser = df.iloc[ser_index_hom,1:].rename(i)  
                    if i == 'tgp' or i== 'to' or i == 'td' or i == 't.d._abierto' or i == 't.d._oculto':
                        ser = ser.set_axis(pd.date_range(start='2001-01-01',periods=len(ser),freq='M')).astype('float')
                    else:
                        ser = ser.set_axis(pd.date_range(start='2001-01-01',periods=len(ser),freq='M')).astype('float')*1000
                    series_hombres[i] = ser

                #Mujeres
                ser_index_muj = df[df.applymap(lambda x: str(x).lower().replace(' ','_') == i)].dropna(how='all',axis=0).index[2]

                if ser_index_muj:
                    ser = df.iloc[ser_index_muj,1:].rename(i)  
                    if i == 'tgp' or i== 'to' or i == 'td' or i == 't.d._abierto' or i == 't.d._oculto':
                        ser = ser.set_axis(pd.date_range(start='2001-01-01',periods=len(ser),freq='M')).astype('float')
                    else:
                        ser = ser.set_axis(pd.date_range(start='2001-01-01',periods=len(ser),freq='M')).astype('float')*1000
                    series_mujeres[i] = ser

    series_tnac.to_csv(path+"\sexo_tnac.csv",sep=';',decimal=',')
    series_hombres.to_csv(path+"\sexo_hombres.csv",sep=';',decimal=',')
    series_hombres.to_csv(path+"\sexo_mujeres.csv",sep=';',decimal=',')
    os.remove(path+"\{}".format(dane_sexo_nombre))
