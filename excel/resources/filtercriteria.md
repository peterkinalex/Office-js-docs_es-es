# Objeto FilterCriteria (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS, Office 2016_

Representa los criterios de filtrado que se aplican a una columna.

## Properties

| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|color|string|Cadena de color HTML que se usa para filtrar las celdas. Se usa con el filtrado de "cellColor" y "fontColor".|
|criterion1|string|Primer criterio usado para filtrar los datos. Se usa como un operador en el caso del filtrado "personalizado".|
|criterion2|string|Segundo criterio usado para filtrar los datos. Solo se usa como un operador en el caso del filtrado "personalizado".|
|dynamicCriteria|string|Criterios dinámicos del conjunto Excel.DynamicFilterCriteria que se van a aplicar a esta columna. Se usa con el filtrado "dinámico". Los valores posibles son: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|
|filterOn|string|Propiedad usada por el filtro para determinar si los valores deben permanecer visibles. Los valores posibles son: 	BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom |
|values|object[]|Conjunto de valores que se van a usar como parte del filtrado de "valores".|

## Relaciones
| Relación | Tipo|Descripción|
|:---------------|:--------|:----------|
|icono|[Icono](icon.md)|Icono usado para filtrar las celdas. Se usa con el filtrado de "icono".|
|operador|FilterOperator|Operador usado para combinar el criterio 1 y 2 cuando se usa el filtrado "personalizado".|

## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método


### load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### Sintaxis
```js
object.load(param);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### Valores devueltos
void

