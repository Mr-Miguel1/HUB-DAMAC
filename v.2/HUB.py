import pandas as pd
import numpy as np
import time
from indicadores import indicadores_xarea
from scrap import scraping_BR,scraping_DANE
from excel import guardar_excel
from limpieza import clean_informalidad


class HUB_DAMAC():

    

    class mercado_laboral():
        
        descripcion_laboral = """
                    La clase mercado_laboral() contiene los idnciadores del mercado laboral usado por la DAMAC
                    se encuentran pre-programados dado que se usa un número limitado de ellos, pero también se pueden añadir otros más 
                    adelante
                    
                    Métodos:
                    
                    -------
                    
                    indicadores():
                                    devuelve una lista con los indicadores del área presentes en la página web del Banco de la república
                    
                    tasa_de_desempleo()
                    tasa_de_ocupación()
                    tasa_de_participación()
                    
                                    devuelve un archivo .csv del indicador respectivo
                    otro_indicador()
                                    
                                    si desea agregar otro indicador puede hacerlo con este método
                    ---------
                    
                    Parámetros:
                    
                    --------
                    
                    carpeta: 
                            Carpeta en la que desea almacenar los archivos, se recomienda que sea creada en un directorio específico 
                            antes de ejecutar la función, ejemplo
                            
                                r"C:\Escritorio\Mercado Laboral" ó
                                
                                "C:/Escritorio/Mercado laboral"
                    indicador:
                            Si opta por usar el método otro_indicador() se le pide que especifique el nombre del indicador que desea
                            actualizar y descargar, el indicador debe estar escrito correctamente, si desconoce el nombre del indicador
                            lo encuentra usando el método indicadores()
                    
                    NOTA: Se recomienda tener el cortafuegos desactivado o darle permisos de administrador para tener acceso a la red
                        pública
                        """
        indicadores_laboral  = pd.DataFrame({
            'Indicador':['Tasa de desempleo','Tasa de ocupación','Tasa global de participación','Informalidad'],
            
            'Frecuencia': ['Mensual','Mensual','Mensual','Trimestral'],
            
            'Fuente': ['https://totoro.banrep.gov.co/estadisticas-economicas/','https://totoro.banrep.gov.co/estadisticas-economicas/',
            'https://totoro.banrep.gov.co/estadisticas-economicas/','https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-informal-y-seguridad-social']
            }) 

        def indicadores_mercado_laboral_BR(self):
            self.indicadores = indicadores_xarea(0)
            return self.indicadores

        
        def actualizar(self,carpeta,actualizar_todo = False,indicadores='',excel=False,hipervinculos=False):

            fuente_laboral = pd.DataFrame({
            'Indicador':['Tasa de desempleo','Tasa de ocupación','Tasa global de participación','Informalidad'],
            'Frecuencia': ['Mensual','Mensual','Mensual','Trimestral'],
            'Fuente': ['https://totoro.banrep.gov.co/estadisticas-economicas/','https://totoro.banrep.gov.co/estadisticas-economicas/',
            'https://totoro.banrep.gov.co/estadisticas-economicas/','https://www.dane.gov.co/index.php/estadisticas-por-tema/mercado-laboral/empleo-informal-y-seguridad-social']
            })
            
            lista_indicadores_BR = ['Tasa de desempleo','Tasa de ocupación','Tasa global de participación','Informalidad']
            # lista_inficadores_DANE = ['']

            if actualizar_todo:
                for i in lista_indicadores_BR:
                    if i != 'Informalidad':
                        try:
                            self.i = scraping_BR(0,indicador=i,path=carpeta)
                        except:
                            continue
                    elif i == 'Informalidad':
                        self.informalidad = scraping_DANE().scraping_dane_mercado_laboral(path=carpeta)
            else:
                for i in indicadores:
                    if i != 'Informalidad':
                        try:
                            self.i = scraping_BR(0,indicador=i,path=carpeta)
                        except:
                            continue
                    elif i == 'Informalidad':
                        self.informalidad = scraping_DANE().scraping_dane_mercado_laboral(path=carpeta)
                        
            try:
                clean_informalidad(path=carpeta)
            except:
                pass

            if excel:
                self.excel = guardar_excel(Fuente=fuente_laboral,carpeta_origen=carpeta,carpeta_destino=carpeta,nombre_archivo='hub_laboral',hyperlinks=hipervinculos)
            else:
                print('Esta bien no genero el excel')

        
    class PIB():

        descripcion_pib =  """
                    La clase PIB() contiene los idnciadores del PIB usado por la DAMAC se encuentran pre-programados dado que
                    se usa un número limitado de ellos, pero también se pueden añadir otros más adelante
                    
                    Métodos:
                    
                    -------
                    
                    indicadores():
                                    devuelve una lista con los indicadores del área presentes en la página web del Banco de la república

                    
                                    devuelve un archivo .csv del indicador respectivo
                                    
                    otro_indicador()
                                    
                                    si desea agregar otro indicador puede hacerlo con este método
                    ---------
                    
                    Parámetros:
                    
                    --------
                    
                    carpeta: 
                            Carpeta en la que desea almacenar los archivos, se recomienda que sea creada en un directorio específico 
                            antes de ejecutar la función, ejemplo
                            
                                r"C:\Escritorio\Mercado Laboral" ó
                                
                                "C:/Escritorio/Mercado laboral"
                    indicador:
                            Si opta por usar el método otro_indicador() se le pide que especifique el nombre del indicador que desea
                            actualizar y descargar, el indicador debe estar escrito correctamente, si desconoce el nombre del indicador
                            lo encuentra usando el método indicadores()
                    
                    NOTA: Se recomienda tener el cortafuegos desactivado o darle permisos de administrador para tener acceso a la red
                        pública
                    """
        indicadores_pib  = pd.DataFrame({
            'Indicador':['Consumo final, nominal','Consumo final, real','Crecimiento PIB nominal',
                         'Crecimiento PIB nominal, ajuste estacional','Crecimiento PIB real',
                         'Crecimiento PIB real, ajuste estacional','Exportaciones, nominal','Exportaciones, real',
                         'Formación bruta de capital, nominal','Formación bruta de capital, real','Importaciones, nominal',
                         'Importaciones, real','Producto Interno Bruto (PIB) nominal',
                         'Producto Interno Bruto (PIB) nominal, ajuste estacional','Producto Interno Bruto (PIB) real',
                         'Producto Interno Bruto (PIB) real, ajuste estacional'],
            'Frecuencia':['Trimestral','Trimestral','Trimestral','Trimestral','Trimestral',
                          'Trimestral','Trimestral','Trimestral','Trimestral','Trimestral',
                          'Trimestral','Trimestral','Trimestral','Trimestral','Trimestral','Trimestral'],
            'Fuente':['https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/','https://totoro.banrep.gov.co/estadisticas-economicas/',
            'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
            'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/']
            })

        def indicadores_pib_BR(self):
            self.indicadores = indicadores_xarea(1)
            return self.indicadores
        
        #Al agregar un nuevo indicador, es recomendable reiniciar el programa y volver a entrar.
        
        def actualizar(self,carpeta,actualizar_todo = False,indicadores='',excel=False,hipervinculos=False):

            fuente_pib = pd.DataFrame({
            'Indicador':['Consumo final, nominal','Consumo final, real','Crecimiento PIB nominal',
                         'Crecimiento PIB nominal, ajuste estacional','Crecimiento PIB real',
                         'Crecimiento PIB real, ajuste estacional','Exportaciones, nominal','Exportaciones, real',
                         'Formación bruta de capital, nominal','Formación bruta de capital, real','Importaciones, nominal',
                         'Importaciones, real','Producto Interno Bruto (PIB) nominal',
                         'Producto Interno Bruto (PIB) nominal, ajuste estacional','Producto Interno Bruto (PIB) real',
                         'Producto Interno Bruto (PIB) real, ajuste estacional'],
            'Frecuencia':['Trimestral','Trimestral','Trimestral','Trimestral','Trimestral',
                          'Trimestral','Trimestral','Trimestral','Trimestral','Trimestral',
                          'Trimestral','Trimestral','Trimestral','Trimestral','Trimestral','Trimestral'],
            'Fuente':['https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/','https://totoro.banrep.gov.co/estadisticas-economicas/',
            'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
            'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/',
                      'https://totoro.banrep.gov.co/estadisticas-economicas/']
            })
            
            lista_indicadores_BR = ['Consumo final, nominal','Consumo final, real','Crecimiento PIB nominal',
                         'Crecimiento PIB nominal, ajuste estacional','Crecimiento PIB real',
                         'Crecimiento PIB real, ajuste estacional','Exportaciones, nominal','Exportaciones, real',
                         'Formación bruta de capital, nominal','Formación bruta de capital, real','Importaciones, nominal',
                         'Importaciones, real','Producto Interno Bruto (PIB) nominal',
                         'Producto Interno Bruto (PIB) nominal, ajuste estacional','Producto Interno Bruto (PIB) real',
                         'Producto Interno Bruto (PIB) real, ajuste estacional']
            # lista_inficadores_DANE = ['']
            

            if actualizar_todo:
                for i in fuente_pib['Indicador']:
                    if i != 'Informalidad':
                        try:
                            self.i = scraping_BR(1,indicador=i,path=carpeta)
                        except:
                            continue
                    elif i == 'Informalidad':
                        self.informalidad = scraping_DANE().scraping_dane_mercado_laboral(path=carpeta)
            else:
                for i in indicadores:
                    if i != 'Informalidad':
                        try:
                            self.i = scraping_BR(1,indicador=i,path=carpeta)
                        except:
                            continue
                    elif i == 'Informalidad':
                        self.informalidad = scraping_DANE().scraping_dane_mercado_laboral(path=carpeta)
                        
            try:
                clean_informalidad(path=carpeta)
            except:
                pass

            if excel:
                self.excel = guardar_excel(Fuente=fuente_pib,carpeta_origen=carpeta,carpeta_destino=carpeta,nombre_archivo='hub_pib',hyperlinks=hipervinculos)
            else:
                print('Esta bien no genero el excel')
                    
        


