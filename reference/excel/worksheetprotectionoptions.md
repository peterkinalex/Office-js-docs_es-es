# <a name="worksheetprotectionoptions-object-(javascript-api-for-excel)"></a>Objeto WorksheetProtectionOptions (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS, Office 2016_

Representa las opciones de protección de hoja.

## <a name="properties"></a>Properties

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|allowAutoFilter|bool|Representa la opción de protección de la hoja de cálculo para permitir usar la característica de filtro automático.|
|allowDeleteColumns|bool|Representa la opción de protección de la hoja de cálculo para permitir eliminar columnas.|
|allowDeleteRows|bool|Representa la opción de protección de la hoja de cálculo para permitir eliminar filas.|
|allowFormatCells|bool|Representa la opción de protección de la hoja de cálculo para permitir aplicar formato a celdas.|
|allowFormatColumns|bool|Representa la opción de protección de la hoja de cálculo para permitir aplicar formato a columnas.|
|allowFormatRows|bool|Representa la opción de protección de la hoja de cálculo para permitir aplicar formato a filas.|
|allowInsertColumns|bool|Representa la opción de protección de la hoja de cálculo para permitir insertar columnas.|
|allowInsertHyperlinks|bool|Representa la opción de protección de la hoja de cálculo para permitir insertar hipervínculos.|
|allowInsertRows|bool|Representa la opción de protección de la hoja de cálculo para permitir insertar filas.|
|allowPivotTables|bool|Representa la opción de protección de la hoja de cálculo para permitir usar la característica de tabla dinámica.|
|allowSort|bool|Representa la opción de protección de la hoja de cálculo para permitir usar la característica de ordenación.|

_Consulte los [ejemplos](#examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## <a name="method-details"></a>Detalles del método


### <a name="load(param:-object)"></a>load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
object.load(param);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
Este ejemplo carga las opciones de protección de la hoja de cálculo activa.
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection options: " + worksheet.protection.options);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
