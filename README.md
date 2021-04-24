# HUB-DAMAC

Es un proyecto que surge en la División de Análisis Macroeconómico como necesidad de contar con una base de datos centralizada que recopile los principales indicadores económcios, sociales y financieros de Colombia. El objetivo central es crear un HUB con los datos reportados de diferentes entidades como el DANE, el Banco de la República, el Miniesterio de Hacienda y Crédito Público, las superintendencias, entre otros.

Librerías utilizadas 

> pandas
> numpy
> selenium
> openpyxl

la versión de python es la 3.7.5

El código cuenta con diferentes modulos que tienen funciones específicas, por ejemplo, el script ``scrap.py`` cuenta con la función scraping_informalidad que permite ingresar a la página del dane y descargar el archivo de informalidad actualizado, así mismo los demás scripts tienn funciónes específicas como hacer data cleaning y generar los archivos de excel.

en el script ``HUB.py`` se utilizan todos los demás con el fin de tener una función que permita scrapear, limpias y generar el excel del sector o área correspondiente


## Anotaciones

Este proyecto nace en la facultad de ciencias económicas de la universidad nacional y es una iniciativa de los estudiantes que la conforman, somos conscientes de que nuestra principal área de conocimiento son las ciencias económicas y no contamos con la experiencia suficiente para elaborar proyectos más estructurados que usen la programación orientada a o objetos entre otras cosas, la participación de un o una programadora, ingeniera de sistemas o de software daría un plus en la elobaración del código. 

El objetivo esque sea un proyecto opensource y que las personas que lo deseen tengan acceso al código y a las bases de datos que estarán en formato xlsx o csv

## Presentación de resultados

Se tendrá una presentación en PowerBi de los datos de cada sector, por el momento se cuenta con el sector laboral y contamos con mas de 30 indicadores que dan cuenta de la situación actual