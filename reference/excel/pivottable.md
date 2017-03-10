# <a name="pivottable-object-javascript-api-for-excel"></a>Objeto PivotTable (API de JavaScript para Excel)

Representa una tabla dinámica de Excel.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|name|string|Nombre de la tabla dinámica.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|worksheet|[Worksheet](worksheet.md)|La hoja de cálculo que contiene la tabla dinámica actual. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|[refresh()](#refresh)|void|Actualiza la tabla dinámica.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="refresh"></a>refresh()
Actualiza la tabla dinámica.

#### <a name="syntax"></a>Sintaxis
```js
pivotTableObject.refresh();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void
