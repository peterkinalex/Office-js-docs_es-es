# <a name="filtercriteria-object-javascript-api-for-excel"></a>Objeto FilterCriteria (API de JavaScript para Excel)

Representa los criterios de filtrado que se aplican a una columna.

## <a name="properties"></a>Properties

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|color|string|Cadena de color HTML que se usa para filtrar las celdas. Se usa con el filtrado de "cellColor" y "fontColor".|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion1|string|Primer criterio usado para filtrar los datos. Se usa como un operador en el caso del filtrado "personalizado".|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion2|string|Segundo criterio usado para filtrar los datos. Solo se usa como un operador en el caso del filtrado "personalizado".|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dynamicCriteria|string|Criterios dinámicos del conjunto Excel.DynamicFilterCriteria que se van a aplicar a esta columna. Se usa con el filtrado "dinámico". Los valores posibles son: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|filterOn|string|Propiedad usada por el filtro para determinar si los valores deben permanecer visibles. Los valores posibles son: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|operator|string|Operador usado para combinar el criterio 1 y 2 cuando se usa el filtrado "personalizado". Los valores posibles son: And, Or.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[]|El conjunto de valores que se va a usar como parte del filtrado de "valores".|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|icono|[Icon](icon.md)|Icono usado para filtrar las celdas. Se usa con el filtrado de "icono".|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos
Ninguna

