# <a name="worksheet-object-javascript-api-for-excel"></a>Objeto Worksheet (API de JavaScript para Excel)

Una hoja de cálculo de Excel es una cuadrícula de celdas. Puede contener datos, tablas, gráficos, etc.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|id|string|Devuelve un valor que identifica de forma única la hoja de cálculo de un libro determinado. El valor del identificador permanece igual, incluso cuando se cambia el nombre de la hoja de cálculo o cuando esta se mueve. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Nombre para mostrar de la hoja de cálculo.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|entero|Posición de base cero de la hoja de cálculo dentro del libro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visibility|string|La visibilidad de la hoja de cálculo. Los valores posibles son: Visible, Hidden, VeryHidden.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|charts|[ChartCollection](chartcollection.md)|Devuelve la colección de gráficos que forman parte de la hoja de cálculo. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|names|[NamedItemCollection](nameditemcollection.md)|Colección de nombres en el ámbito de la hoja de cálculo actual. Solo lectura.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|Colección de tablas dinámicas que forman parte de la hoja de cálculo. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[WorksheetProtection](worksheetprotection.md)|Devuelve el objeto de protección de hoja de una hoja de cálculo. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|Colección de tablas que forman parte de la hoja de cálculo. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[activate()](#activate)|void|Activa la hoja de cálculo en la interfaz de usuario de Excel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|Elimina la hoja de cálculo del libro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Obtiene el objeto de rango que contiene la celda en función de los números de fila y columna. La celda puede estar fuera de los límites del rango principal, siempre y cuando permanezca dentro de la cuadrícula de la hoja de cálculo.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|Obtiene el objeto de intervalo especificado por la dirección o el nombre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly: [ApiSet(Version)](#getusedrangevaluesonly-apisetversion)|[Range](range.md)|El rango usado es el rango más pequeño que abarque todas las celdas que tengan asignado un valor o un formato. Si toda la hoja está en blanco, esta función devuelve la celda superior izquierda (es decir, *no* genera un error).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeOrNullObject(valuesOnly: bool)](#getusedrangeornullobjectvaluesonly-bool)|[Range](range.md)|El rango usado es el rango más pequeño que abarque todas las celdas que tengan asignado un valor o un formato. Si toda la hoja está en blanco, esta función devolverá un objeto NULL.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="activate"></a>activate()
Activa la hoja de cálculo en la interfaz de usuario de Excel.

#### <a name="syntax"></a>Sintaxis
```js
worksheetObject.activate();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.activate();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="delete"></a>delete()
Elimina la hoja de cálculo del libro.

#### <a name="syntax"></a>Sintaxis
```js
worksheetObject.delete();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcellrow-number-column-number"></a>getCell(row: number, column: number)
Obtiene el objeto de rango que contiene la celda en función de los números de fila y columna. La celda puede estar fuera de los límites del rango principal, siempre y cuando permanezca dentro de la cuadrícula de la hoja de cálculo.

#### <a name="syntax"></a>Sintaxis
```js
worksheetObject.getCell(row, column);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|row|number|Número de fila de la celda que se va a recuperar. Indizado con cero.|
|column|number|Número de columna de la celda que se va a recuperar. Indexado con cero.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var cell = worksheet.getCell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrangeaddress-string"></a>getRange(address: string)
Obtiene el objeto de intervalo especificado por la dirección o el nombre.

#### <a name="syntax"></a>Sintaxis
```js
worksheetObject.getRange(address);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|address|string|Opcional. Dirección o nombre del intervalo. Si no se especifica, se devuelve todo el intervalo de la hoja de cálculo.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
En el ejemplo siguiente se usa la dirección del intervalo para obtener el objeto de intervalo.

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

En el ejemplo siguiente se usa un intervalo con nombre para obtener el objeto de intervalo.

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

### <a name="getusedrangevaluesonly-apisetversion"></a>getUsedRange(valuesOnly: [ApiSet(Version)
El rango usado es el rango más pequeño que abarque todas las celdas que tengan asignado un valor o un formato. Si toda la hoja está en blanco, esta función devuelve la celda superior izquierda (es decir, *no* genera un error).

#### <a name="syntax"></a>Sintaxis
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|valuesOnly|[ApiSet(Version|Solo tiene en cuenta las celdas con valores como celdas usadas (ignora el formato).|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

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


### <a name="getusedrangeornullobjectvaluesonly-bool"></a>getUsedRangeOrNullObject(valuesOnly: bool)
El rango usado es el rango más pequeño que abarque todas las celdas que tengan asignado un valor o un formato. Si toda la hoja está en blanco, esta función devolverá un objeto NULL.

#### <a name="syntax"></a>Sintaxis
```js
worksheetObject.getUsedRangeOrNullObject(valuesOnly);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|valuesOnly|bool|Opcional. Solo tiene en cuenta las celdas con valores como celdas usadas.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

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
