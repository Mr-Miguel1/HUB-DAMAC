from HUB import HUB_DAMAC

laboral = HUB_DAMAC.mercado_laboral()
laboral.actualizar(carpeta=r"D:\Desktop\Laboral",actualizar_todo=True, excel=True,hipervinculos=True)
# print(laboral.descripcion)
# laboral.fuente_laboral
# guardar_excel(Fuente=laboral.fuente_laboral,carpeta_origen=r"D:\Desktop\Laboral",carpeta_destino=r"D:\Desktop\Laboral", nombre_archivo='hub_laboral',hyperlinks=True)
#laboral.actualizar(carpeta=r"D:\Desktop\Laboral",actualizar_todo=False,indicadores=['Tasa de desempleo','Tasa de ocupación'],excel=True,hipervinculos=True)
# laboral.actualizar(carpeta=r"D:\Desktop\Laboral",actualizar_todo=False,indicadores=['Tasa de desempleo','Informalidad'],excel=True,hipervinculos=True)

