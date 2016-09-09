# Objeto Worksheet (API de JavaScript para Excel)

Una hoja de cálculo de Excel es una cuadrícula de celdas. Puede contener datos, tablas, gráficos, etc.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|id|string|Devuelve un valor que identifica de forma única la hoja de cálculo de un libro determinado. El valor del identificador permanece igual, incluso cuando se cambia el nombre de la hoja de cálculo o cuando esta se mueve. Los valores cambian con cada sesión del archivo que se abre. Solo lectura.|
|name|string|Nombre para mostrar de la hoja de cálculo.|
|position|entero|Posición de base cero de la hoja de cálculo dentro del libro.|
|visibility|string|Visibilidad de la hoja de cálculo. Los valores posibles son: Visible, Hidden, VeryHidden.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|charts|[ChartCollection](chartcollection.md)|Devuelve la colección de gráficos que forman parte de la hoja de cálculo. Solo lectura.|
|protección|[WorksheetProtection](worksheetprotection.md)|Devuelve el objeto de protección de hoja de una hoja de cálculo. Solo lectura.|
|tablas|[TableCollection](tablecollection.md)|Colección de tablas que forman parte de la hoja de cálculo. Solo lectura.|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|Activa la hoja de cálculo en la interfaz de usuario de Excel.|
|[delete()](#delete)|void|Elimina la hoja de cálculo del libro.|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Obtiene el objeto de intervalo que contiene la celda en función de los números de fila y columna. La celda puede estar fuera de los límites del intervalo principal, siempre y cuando permanezca dentro de la cuadrícula de la hoja de cálculo.|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|Obtiene el objeto de intervalo especificado por la dirección o el nombre.|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Range](range.md)|El intervalo usado es el intervalo más pequeño que abarque las celdas que tienen asignado un valor o un formato. Si la hoja de cálculo está en blanco, esta función devolverá la celda superior izquierda.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método


### activate()
Activa la hoja de cálculo en la interfaz de usuario de Excel.

#### Sintaxis
```js
worksheetObject.activate();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.activate();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### delete()
Elimina la hoja de cálculo del libro.

#### Sintaxis
```js
worksheetObject.delete();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.delete();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getCell(row: number, column: number)
Obtiene el objeto de intervalo que contiene la celda en función de los números de fila y columna. La celda puede estar fuera de los límites del intervalo principal, siempre y cuando permanezca dentro de la cuadrícula de la hoja de cálculo.

#### Sintaxis
```js
worksheetObject.getCell(row, column);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|row|number|Número de fila de la celda que se va a recuperar. Indizado con cero.|
|column|number|Número de columna de la celda que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[Range](range.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var cell = worksheet.getCell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getRange(address: string)
Obtiene el objeto de intervalo especificado por la dirección o el nombre.

#### Sintaxis
```js
worksheetObject.getRange(address);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|address|string|Opcional. Dirección o nombre del intervalo. Si no se especifica, se devuelve todo el intervalo de la hoja de cálculo.|

#### Valores devueltos
[Range](range.md)

#### Ejemplos
En este ejemplo se usa la dirección del intervalo para obtener el objeto de intervalo.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

En este ejemplo se usa un intervalo con nombre para obtener el objeto de intervalo.

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeName = 'MyRange';
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getUsedRange(valuesOnly: bool)
El intervalo usado es el intervalo más pequeño que abarque las celdas que tienen asignado un valor o un formato. Si la hoja de cálculo está en blanco, esta función devolverá la celda superior izquierda.

#### Sintaxis
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|valuesOnly|bool|Opcional. Solo tiene en cuenta las celdas con valores como celdas usadas (ignora el formato).|

#### Valores devueltos
[Range](range.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    var usedRange = worksheet.getUsedRange();
    usedRange.load('address');
    return ctx.sync().then(function() {
            console.log(usedRange.address);
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

Obtener las propiedades de la hoja de cálculo en función del nombre de la hoja.

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.load('position')
    return ctx.sync().then(function() {
            console.log(worksheet.position);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Establecer la posición de la hoja de cálculo. 

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.position = 2;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

