# Objeto RangeFormat (API de JavaScript para Excel)

Objeto de formato que encapsula la fuente, el relleno, los bordes, la alineación y otras propiedades del intervalo.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|columnWidth|double|Obtiene o establece el ancho de todas las columnas del rango. Si los anchos de columna no son uniformes, se devolverá nulo.|
|horizontalAlignment|string|Representa la alineación horizontal del objeto especificado. Los valores posibles son: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|
|rowHeight|double|Obtiene o establece el alto de todas las filas del rango. Si los altos de fila no son uniformes, se devolverá nulo.|
|verticalAlignment|string|Representa la alineación vertical del objeto especificado. Los valores posibles son: Top, Center, Bottom, Justify, Distributed.|
|wrapText|bool|Indica que el control de texto de Excel está establecido para ajustar texto en el objeto. Un valor null indica que el intervalo no usa una configuración de ajuste de texto uniforme.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|borders|[RangeBorderCollection](rangebordercollection.md)|Colección de objetos de borde que se aplican al intervalo global seleccionado. Solo lectura.|
|fill|[RangeFill](rangefill.md)|Devuelve el objeto de relleno definido en el intervalo global. Solo lectura.|
|font|[RangeFont](rangefont.md)|Devuelve el objeto de fuente definido en el intervalo global seleccionado. Solo lectura.|
|protección|[FormatProtection](formatprotection.md)|Devuelve el objeto de protección de formato de un rango. Solo lectura.|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[autofitColumns()](#autofitcolumns)|void|Cambia el ancho de las columnas del rango actual para obtener el ajuste perfecto (según los datos actuales de las columnas).|
|[autofitRows()](#autofitrows)|void|Cambia el alto de las filas del rango actual para obtener el ajuste perfecto (según los datos actuales de las columnas).|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método


### autofitColumns()
Cambia el ancho de las columnas del rango actual para obtener el ajuste perfecto (según los datos actuales de las columnas).

#### Sintaxis
```js
rangeFormatObject.autofitColumns();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

### autofitRows()
Cambia el alto de las filas del rango actual para obtener el ajuste perfecto (según los datos actuales de las columnas).

#### Sintaxis
```js
rangeFormatObject.autofitRows();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

### load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### Sintaxis
```js
object.load(param);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### Valores devueltos
void
### Ejemplos de acceso a la propiedad

En este ejemplo se imprimen todas las propiedades de formato de un intervalo. 

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

En el ejemplo siguiente se establecen el nombre de fuente y el color de relleno de un intervalo y se ajusta el texto. 

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
    range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
    range.format.borders('InsideVertical').lineStyle = 'Continuous';
    range.format.borders('EdgeBottom').lineStyle = 'Continuous';
    range.format.borders('EdgeLeft').lineStyle = 'Continuous';
    range.format.borders('EdgeRight').lineStyle = 'Continuous';
    range.format.borders('EdgeTop').lineStyle = 'Continuous';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
