# <a name="sortfield-object-javascript-api-for-excel"></a>Objeto SortField (API de JavaScript para Excel)

Representa una condición en una operación de ordenación.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|ascending|bool|Representa si la ordenación se realiza en orden ascendente.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|Representa el color que es el destino de la condición si la ordenación se realiza según la fuente o el color de celda.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dataOption|string|Representa opciones de ordenación adicionales para este campo. Los valores posibles son: Normal, TextAsNumber.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|key|int|Representa la columna (o fila, según la orientación de ordenación) en que se encuentra la condición. Se representa como un desplazamiento de la primera columna (o fila).|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|sortOn|string|Representa el tipo de ordenación de esta condición. Los valores posibles son: Value, CellColor, FontColor, Icon.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|icono|[Icon](icon.md)|Representa el icono que es el destino de la condición si la ordenación se realiza según el icono de la celda.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos
Ninguna

