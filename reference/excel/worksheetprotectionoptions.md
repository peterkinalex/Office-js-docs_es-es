# <a name="worksheetprotectionoptions-object-javascript-api-for-excel"></a>Objeto WorksheetProtectionOptions (API de JavaScript para Excel)

Representa las opciones de protección de hoja.

## <a name="properties"></a>Properties

| Propiedad     | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|allowAutoFilter|bool|Representa la opción de protección de la hoja de cálculo que permite usar la característica de filtro automático.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowDeleteColumns|bool|Representa la opción de protección de la hoja de cálculo que permite eliminar columnas.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowDeleteRows|bool|Representa la opción de protección de la hoja de cálculo que permite eliminar filas.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatCells|bool|Representa la opción de protección de la hoja de cálculo que permite aplicar formato a celdas.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatColumns|bool|Representa la opción de protección de la hoja de cálculo que permite aplicar formato a columnas.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatRows|bool|Representa la opción de protección de la hoja de cálculo que permite aplicar formato a filas.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertColumns|bool|Representa la opción de protección de la hoja de cálculo que permite insertar columnas.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertHyperlinks|bool|Representa la opción de protección de la hoja de cálculo que permite insertar hipervínculos.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertRows|bool|Representa la opción de protección de la hoja de cálculo que permite insertar filas.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowPivotTables|bool|Representa la opción de protección de la hoja de cálculo que permite usar la característica de tabla dinámica.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowSort|bool|Representa la opción de protección de la hoja de cálculo que permite usar la característica de ordenación.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy que se ha creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="loadparam-object"></a>load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
object.load(param);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
void
