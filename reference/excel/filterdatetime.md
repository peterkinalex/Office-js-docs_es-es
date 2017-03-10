# <a name="filterdatetime-object-javascript-api-for-excel"></a>Objeto FilterDatetime (API de JavaScript para Excel)

Representa cómo se filtra una fecha cuando se filtran valores.

## <a name="properties"></a>Properties

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|date|string|La fecha en formato ISO8601 usada para filtrar los datos.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|specificity|string|El grado de especificidad de la fecha que se usará para mantener datos. Por ejemplo, si la fecha es 02-04-2005 y la especificidad se establece en "mes", la operación de filtrado conservará todas las filas con fecha de abril de 2005. Los valores posibles son: Year, Monday, Day, Hour, Minute, Second.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos
Ninguna

