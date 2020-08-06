# stock_values

stock_values obtiene los valores de los Cedears que se encuentren en una planilla de Excel. Fue creado con el propósito de automatizar la búsqueda de los valores de cada Cedear. Permite también agregar alguna acción que no se encuentre en la planilla.

## Requerimientos

Se necesita selenium y openpyxl. Tener Firefox instalado también es recomendable.

Por último, es necesario tener en el mismo directorio una hoja de Excel con las fechas hasta el día actual en la columna A, y a partir de la segunda fila. No es necesario tener el nombre de las acciones.

## Uso

El programa va a buscar por defecto las acciones que se encuentren en la planilla. Luego va a preguntar si se quiere agregar alguna más. Es importante saber siempre en qué bolsa cotiza, de lo contrario va a fallar.
