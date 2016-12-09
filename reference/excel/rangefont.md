# <a name="rangefont-object-javascript-api-for-excel"></a>Objeto RangeFont (API de JavaScript para Excel)

Este objeto representa los atributos de fuente (nombre de fuente, tamaño de fuente, color, etc.) de un objeto.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|bold|bool|Representa el estado de negrita de la fuente.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|Representación del código de color HTML del color del texto. Por ejemplo, #FF0000 representa el rojo.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|Representa el estado de cursiva de la fuente.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Nombre de fuente (por ejemplo, "Calibri")|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|size|Double|Tamaño de fuente.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|underline|string|Tipo de subrayado aplicado a la fuente. Los valores posibles son: None, Single, Double, SingleAccountant, DoubleAccountant.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var rangeFont = range.format.font;
    rangeFont.load('name');
    return ctx.sync().then(function() {
        console.log(rangeFont.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
En el ejemplo siguiente se establece el nombre de la fuente. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.font.name = 'Times New Roman';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```