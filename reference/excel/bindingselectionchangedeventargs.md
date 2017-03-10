# <a name="bindingselectionchangedeventargs-object-javascript-api-for-excel"></a>Objeto BindingSelectionChangedEventArgs (API de JavaScript para Excel)

Proporciona información sobre el enlace que ha generado el evento SelectionChanged.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|columnCount|entero|Obtiene la cantidad de columnas seleccionadas.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|entero|Obtiene la cantidad de filas seleccionadas.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startColumn|entero|Obtiene el índice de la primera columna de la selección (de base cero).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startRow|entero|Obtiene el índice de la primera fila de la selección (de base cero).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|enlace|[Binding](binding.md)|Obtiene un objeto Binding que representa el enlace que ha generado el evento SelectionChanged.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos
Ninguna

