# Overview of Project

The project is based on the analysis of price and volume data of a number of stocks in the green energy sector in order to determine which stocks offer a positive return and which do not, so that an investment decision can be made.

As a starting point, an Excel file was received with information on a group of 12 pre-selected stocks.  The Excel file contains 8 columns and 3013 rows.  It was not necessary to work with all the columns, but only with column 1 containing the stock name code, also known as ticker, column 6 containing the daily closing price of the stock and column 8 containing the daily amount of volume traded on the stock exchange for the different stocks.

### Purpose of the analysis

Este proyecto tiene dos partes.

#### Parte 1
La finalidad de la primera parte de este proyecto es la automatización del análisis de las acciones, mediante la creación de un código que lea los valores con el precio de la acción al principio del año y al cierre del año y saque el rendimiento de la acción en forma de porcentaje.  El código también debe reportar el volumen total transado por cada acción.  Al finalizar de generar la información indicada, el código debe crear una tabla con los resultados, formatear las columnas de dicha tabla y resaltar con color verde aquellas que obtuvieron ganancia y con color rojo las acciones que tuvieron un desempeño negativo.

El código original trabaja en base a dos loops anidados.  Un loop recorre los tickers de las 12 acciones y para cada ticker realiza otro loop que recorre toda la hoja de Excel y recolecta la información sobre el precio inicial, precio final y volumen de cada uno de los tickers.

#### Parte 2
En la segunda parte, se busca mejorar el tiempo de ejecución del código y se propone cambiar (Refactor) el código original.  El cambio consiste en hacer un loop que recorra toda la hoja de Excel y va llevando el total acumulado para las acciones.  Además, cada vez que se produce un cambio en el nombre del ticker, registra el precio inicial y final del mismo.

La idea es comparar el tiempo de ejecución del método original descrito en la parte 1 contra el tiempo de ejecución del código refactored descrito en la parte 2 y ver si existe diferencia al hacer el refactoring.

#### Parte 3
Esta parte fue desarrollada por iniciativa propia con la finalidad de ver si se podía mejorar aún mas el código refactored.

En el tercer método se recorren los datos de la hoja de Excel una sola vez y se van guardando en un array los números de fila en los que sucede el cambio de un ticker a otroy se saca el dato del volumen acumulado para cada ticker.  A estas filas se les denomina break points.  Una vez determinados los break points, se hace un llamado a las celdas de Excel que contienen los datos del precio de inicial y final de cada acción y se construye la tabla con los resultados, dandole el mismo for,mato de colores verde y rojo explicado en la sección anterior.




## Results

## Conclusions



