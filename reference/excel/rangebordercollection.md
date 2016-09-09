# Objeto RangeBorderCollection (API de JavaScript para Excel)

Representa los objetos de borde que componen el borde del intervalo.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|count|int|Número de objetos de borde de la colección. Solo lectura.|
|Items|[RangeBorder[]](rangeborder.md)|Colección de objetos rangeBorder. Solo lectura.|

_Consulte los [jemplos](#jemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[getItem(index: string)](#getitemindex-string)|[RangeBorder](rangeborder.md)|Obtiene un objeto de borde mediante su nombre.|
|[getItemAt(index: number)](#getitematindex-number)|[RangeBorder](rangeborder.md)|Obtiene un objeto de borde mediante su índice.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método


### getItem(index: string)
Obtiene un objeto de borde mediante su nombre. 

#### Sintaxis
```js
rangeBorderCollectionObject.getItem(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|string|Valor de índice del objeto de borde que va a recuperarse.  Los valores posibles son: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight, InsideVertical, InsideHorizontal, DiagonalDown, DiagonalUp.|

#### Valores devueltos
[RangeBorder](rangeborder.md)

#### Ejemplos
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var borderName = 'EdgeTop';
    var border = range.format.borders.getItem(borderName);
    border.load('style');
    return ctx.sync().then(function() {
            console.log(border.style);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


#### Ejemplos
```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var border = ctx.workbook.borders.getItemAt(0);
    border.load('sideIndex');
    return ctx.sync().then(function() {
            console.log(border.sideIndex);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItemAt(index: number)
Obtiene un objeto de borde mediante su índice.

#### Sintaxis
```js
rangeBorderCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[RangeBorder](rangeborder.md)

#### Ejemplos
```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var border = ctx.workbook.borders.getItemAt(0);
    border.load('sideIndex');
    return ctx.sync().then(function() {
            console.log(border.sideIndex);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


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

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var borders = range.format.borders;
    borders.load('items');
    return ctx.sync().then(function() {
        console.log(borders.count);
        for (var i = 0; i < borders.items.length; i++)
        {
            console.log(borders.items[i].sideIndex);
        }
    });
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
    var rangeAddress = "A1:F8";
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
