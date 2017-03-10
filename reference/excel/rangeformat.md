# <a name="rangeformat-object-javascript-api-for-excel"></a>Objeto RangeFormat (API de JavaScript para Excel)

Objeto de formato que encapsula la fuente, el relleno, los bordes, la alineación y otras propiedades del rango.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|columnWidth|double|Obtiene o establece el ancho de todas las columnas del rango. Si los anchos de columna no son uniformes, se devolverá null.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalAlignment|string|Representa la alineación horizontal del objeto especificado. Los valores posibles son: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowHeight|double|Obtiene o establece el alto de todas las filas del rango. Si los altos de fila no son uniformes, se devolverá null.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|verticalAlignment|string|Representa la alineación vertical del objeto especificado. Los valores posibles son: Top, Center, Bottom, Justify, Distributed.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|wrapText|bool|Indica si Excel ajusta el texto del objeto. Un valor null indica que el intervalo no tiene una configuración de ajuste uniforme.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|borders|[RangeBorderCollection](rangebordercollection.md)|Colección de objetos de borde que se aplica al rango global seleccionado. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|fill|[RangeFill](rangefill.md)|Devuelve el objeto de relleno definido en el rango global. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|font|[RangeFont](rangefont.md)|Devuelve el objeto de fuente definido en el rango global. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[FormatProtection](formatprotection.md)|Devuelve el objeto de protección de formato de un rango. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[autofitColumns()](#autofitcolumns)|void|Cambia el ancho de las columnas del intervalo actual para obtener el ajuste perfecto (según los datos actuales de las columnas).|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[autofitRows()](#autofitrows)|void|Cambia el alto de las filas del intervalo actual para obtener el ajuste perfecto (según los datos actuales de las columnas).|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="autofitcolumns"></a>autofitColumns()
Cambia el ancho de las columnas del rango actual para obtener el ajuste perfecto (según los datos actuales de las columnas).

#### <a name="syntax"></a>Sintaxis
```js
rangeFormatObject.autofitColumns();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

### <a name="autofitrows"></a>autofitRows()
Cambia el alto de las filas del rango actual para obtener el ajuste perfecto (según los datos actuales de las columnas).

#### <a name="syntax"></a>Sintaxis
```js
rangeFormatObject.autofitRows();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

El ejemplo siguiente selecciona todas las propiedades de formato del intervalo. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load(["format/*", "format/fill", "format/borders", "format/font"]);
    return ctx.sync().then(function() {
        console.log(range.format.wrapText);
        console.log(range.format.fill.color);
        console.log(range.format.font.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

En el ejemplo siguiente se establecen el nombre de fuente y el color de relleno y se ajusta el texto. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.wrapText = true;
    range.format.font.name = 'Times New Roman';
    range.format.fill.color = '0000FF';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

En el ejemplo siguiente se agrega un borde de cuadrícula alrededor del intervalo.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
    range.format.borders.getItem('InsideVertical').style = 'Continuous';
    range.format.borders.getItem('EdgeBottom').style = 'Continuous';
    range.format.borders.getItem('EdgeLeft').style = 'Continuous';
    range.format.borders.getItem('EdgeRight').style = 'Continuous';
    range.format.borders.getItem('EdgeTop').style = 'Continuous';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```