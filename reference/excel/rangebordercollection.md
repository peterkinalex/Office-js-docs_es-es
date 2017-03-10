# <a name="rangebordercollection-object-javascript-api-for-excel"></a>Objeto RangeBorderCollection (API de JavaScript para Excel)

Representa los objetos de borde que componen el borde del rango.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|count|int|Número de objetos de borde de la colección. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[RangeBorder[]](rangeborder.md)|Colección de objetos rangeBorder. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[getItem(index: string)](#getitemindex-string)|[RangeBorder](rangeborder.md)|Obtiene un objeto de borde mediante su nombre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[RangeBorder](rangeborder.md)|Obtiene un objeto de borde mediante su índice.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="getitemindex-string"></a>getItem(index: string)
Obtiene un objeto de borde mediante su nombre.

#### <a name="syntax"></a>Sintaxis
```js
rangeBorderCollectionObject.getItem(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|string|Valor de índice del objeto de borde que va a recuperarse.  Los valores posibles son: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight, InsideVertical, InsideHorizontal, DiagonalDown, DiagonalUp.|

#### <a name="returns"></a>Valores devueltos
[RangeBorder](rangeborder.md)

#### <a name="examples"></a>Ejemplos
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


#### <a name="examples"></a>Ejemplos
```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var border = range.format.borders.getItemAt(0);
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


### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtiene un objeto de borde mediante su índice.

#### <a name="syntax"></a>Sintaxis
```js
rangeBorderCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[RangeBorder](rangeborder.md)

#### <a name="examples"></a>Ejemplos
```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var border = range.format.borders.getItemAt(0);
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

### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var borders = range.format.borders;
    border.load('items');
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