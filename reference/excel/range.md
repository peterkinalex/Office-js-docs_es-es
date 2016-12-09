# <a name="range-object-javascript-api-for-excel"></a>Objeto Range (API de JavaScript para Excel)

Range representa un conjunto de una o más celdas contiguas, como una celda, una fila, una columna, un bloque de celdas, etc.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|address|string|Representa la referencia de intervalo en estilo A1. El valor de dirección contendrá la referencia de hoja (por ejemplo, Sheet1!A1:B4). Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|addressLocal|string|Representa la referencia del intervalo especificado en el idioma del usuario. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|cellCount|entero|Número de celdas del intervalo. Esta API devolverá -1 si el recuento de celdas supera 2^31-1 (2 147 483 647). Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|entero|Representa el número total de columnas del intervalo. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnHidden|bool|Representa si todas las columnas del intervalo actual están ocultas.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|columnIndex|entero|Representa el número de columna de la primera celda del intervalo. Indizado con cero. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|Representa la fórmula en notación de estilo A1.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|Representa la fórmula en notación de estilo A1, en el idioma del usuario y en la configuración regional del formato numérico. Por ejemplo, la fórmula "=SUM(A1, 1.5)" en inglés se convertiría en "=SUMME(A1; 1,5)" en alemán.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|Representa la fórmula en notación de estilo R1C1.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|hidden|bool|Representa si todas las celdas del rango actual están ocultas. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|object[][]|Representa el código de formato numérico de Excel para la celda especificada.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|Devuelve el número total de filas del intervalo. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowHidden|bool|Representa si todas las filas del intervalo actual están ocultas.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|rowIndex|entero|Devuelve el número de fila de la primera celda del intervalo. Indizado con cero. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|text|.|Valores de texto del rango especificado. El valor Text no dependerá del ancho de la celda. La sustitución del signo # que tiene lugar en la interfaz de usuario de Excel no afectará al valor de texto devuelto por la API. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|Representa el tipo de datos de cada celda. Solo lectura. Los valores posibles son: Unknown, Empty, String, Integer, Double, Boolean, Error.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|Representa los valores sin formato del rango especificado. Los datos devueltos pueden ser de tipo string, number o boolean. La celda que contenga un error devolverá la cadena de error.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|format|[RangeFormat](rangeformat.md)|Devuelve un objeto de formato que encapsula la fuente, el relleno, los bordes, la alineación y otras propiedades del intervalo. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sort|[RangeSort](rangesort.md)|Representa la ordenación del intervalo del intervalo actual. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|Hoja de cálculo que contiene el rango actual. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[clear(applyTo: string)](#clearapplyto-string)|nulo|Borra valores de intervalo, formato, relleno, borde, etc.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete(shift: string)](#deleteshift-string)|nulo|Elimina las celdas asociadas al intervalo.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getBoundingRect(anotherRange: Range o string)](#getboundingrectanotherrange-range-or-string)|[Range](range.md)|Obtiene el objeto de rango más pequeño que abarca los rangos especificados. Por ejemplo, el valor GetBoundingRect de "B2:C5" y "D10:E15" es "B2:E16".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Obtiene el objeto de rango que contiene la celda en función de los números de fila y columna. La celda puede estar fuera de los límites del rango principal, siempre y cuando permanezca dentro de la cuadrícula de la hoja de cálculo. La celda devuelta se ubica con respecto a la celda superior izquierda del rango.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|Obtiene una columna contenida en el intervalo.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsAfter(count: number)](#getcolumnsaftercount-number)|[Range](range.md)|Obtiene un número determinado de columnas a la derecha del objeto Range actual.|[1.2, 1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsBefore(count: number)](#getcolumnsbeforecount-number)|[Range](range.md)|Obtiene un número determinado de columnas a la izquierda del objeto Range actual.|[1.2, 1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireColumn()](#getentirecolumn)|[Range](range.md)|Obtiene un objeto que representa toda la columna del intervalo.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireRow()](#getentirerow)|[Range](range.md)|Obtiene un objeto que representa toda la fila del intervalo.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersection(anotherRange: Range o string)](#getintersectionanotherrange-range-or-string)|[Range](range.md)|Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersectionOrNull(anotherRange: Range or string)](#getintersectionornullanotherrange-range-or-string)|[Range](range.md)|Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados. Si no se encuentra ninguna intersección, se devolverá un objeto NULL.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastCell()](#getlastcell)|[Range](range.md)|Obtiene la última celda del intervalo. Por ejemplo, la última celda de "B2:D5" es "D5".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastColumn()](#getlastcolumn)|[Range](range.md)|Obtiene la última columna del intervalo. Por ejemplo, la última columna de "B2:D5" es "D2:D5".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastRow()](#getlastrow)|[Range](range.md)|Obtiene la última fila del intervalo. Por ejemplo, la última fila de "B2:D5" es "B5:D5".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|Obtiene un objeto que representa un rango desplazado con respecto al rango especificado. La dimensión del rango devuelto coincidirá con este rango. Si el rango resultante se fuerza fuera de los límites de la cuadrícula de la hoja de cálculo, se producirá una excepción.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getResizedRange(deltaRows: number, deltaColumns: number)](#getresizedrangedeltarows-number-deltacolumns-number)|[Range](range.md)|Obtiene un objeto Range similar al objeto Range actual, pero con su esquina inferior derecha expandida (o contraída) mediante un número de filas y columnas.|[1.2, 1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|Obtiene una fila contenida en el intervalo.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsAbove(count: number)](#getrowsabovecount-number)|[Range](range.md)|Obtiene un número determinado de filas encima del objeto Range actual.|[1.2, 1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsBelow(count: number)](#getrowsbelowcount-number)|[Range](range.md)|Obtiene un número determinado de filas debajo del objeto Range actual.|[1.2, 1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly)](#getusedrangevaluesonly)|[Range](range.md)|Devuelve el intervalo usado del objeto de intervalo especificado.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getVisibleView()](#getvisibleview)|[RangeView](rangeview.md)|Representa las filas visibles del intervalo actual.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|Inserta una celda o un rango de celdas en la hoja de cálculo en lugar de este rango y desplaza las demás celdas para crear espacio. Devuelve un objeto Range en el espacio que queda en blanco.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy que se ha creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[merge(across: bool)](#mergeacross-bool)|void|Combina las celdas del intervalo en una región de la hoja de cálculo.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[select()](#select)|void|Selecciona el intervalo especificado en la interfaz de usuario de Excel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[unmerge()](#unmerge)|void|Separa las celdas del intervalo en celdas independientes.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="clearapplyto-string"></a>clear(applyTo: string)
Borra valores de intervalo, formato, relleno, borde, etc.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.clear(applyTo);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|applyTo|string|Opcional. Determina el tipo de acción de borrado. Los valores posibles son: Opción predeterminada `All`, `Formats` ,`Contents` |

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos

En el ejemplo siguiente se borra el formato y el contenido del intervalo. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="deleteshift-string"></a>delete(shift: string)
Elimina las celdas asociadas al intervalo.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.delete(shift);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|Shift|string|Especifica hacia dónde se desplazarán las celdas.  Los valores posibles son: Up, Left|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getboundingrectanotherrange-range-or-string"></a>getBoundingRect(anotherRange: Range or string)
Obtiene el objeto de intervalo más pequeño que abarca los intervalos especificados. Por ejemplo, el valor GetBoundingRect de "B2:C5" y "D10:E15" es "B2:E16".

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getBoundingRect(anotherRange);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|anotherRange|Range o string|Objeto o dirección de intervalo o nombre de intervalo.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var range = range.getBoundingRect("G4:H8");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // Prints Sheet1!D4:H8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcellrow-number-column-number"></a>getCell(row: number, column: number)
Obtiene el objeto de intervalo que contiene la celda en función de los números de fila y columna. La celda puede estar fuera de los límites del intervalo principal, siempre y cuando permanezca dentro de la cuadrícula de la hoja de cálculo. La celda devuelta se ubica con respecto a la celda superior izquierda del intervalo.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getCell(row, column);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|row|number|Número de fila de la celda que se va a recuperar. Indizado con cero.|
|column|number|Número de columna de la celda que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var cell = range.cell(0,0);
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


### <a name="getcolumncolumn-number"></a>getColumn(column: number)
Obtiene una columna contenida en el intervalo.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getColumn(column);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|column|number|Número de columna del intervalo que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet19";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getColumn(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!B1:B8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcolumnsaftercount-number"></a>getColumnsAfter(count: number)
Obtiene un número determinado de columnas a la derecha del objeto Range actual.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getColumnsAfter(count);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|count|number|Opcional. El número de columnas que se va a incluir en el intervalo resultante. En general, use un número positivo para crear un intervalo fuera del intervalo actual. También puede usar un número negativo para crear un intervalo dentro del intervalo actual. El valor predeterminado es 1.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

### <a name="getcolumnsbeforecount-number"></a>getColumnsBefore(count: number)
Obtiene un número determinado de columnas a la izquierda del objeto Range actual.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getColumnsBefore(count);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|count|number|Opcional. El número de columnas que se va a incluir en el intervalo resultante. En general, use un número positivo para crear un intervalo fuera del intervalo actual. También puede usar un número negativo para crear un intervalo dentro del intervalo actual. El valor predeterminado es 1.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

### <a name="getentirecolumn"></a>getEntireColumn()
Obtiene un objeto que representa toda la columna del intervalo.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getEntireColumn();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

Nota: Las propiedades de cuadrícula del intervalo (values, numberFormat, formulas) contienen `null`, dado que el intervalo en cuestión está desvinculado.

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeEC = range.getEntireColumn();
    rangeEC.load('address');
    return ctx.sync().then(function() {
        console.log(rangeEC.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getentirerow"></a>getEntireRow()
Obtiene un objeto que representa toda la fila del intervalo.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getEntireRow();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "D:F"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeER = range.getEntireRow();
    rangeER.load('address');
    return ctx.sync().then(function() {
        console.log(rangeER.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Las propiedades de cuadrícula del intervalo (values, numberFormat, formulas) contienen `null`, dado que el intervalo en cuestión está desvinculado.


### <a name="getintersectionanotherrange-range-or-string"></a>getIntersection(anotherRange: Range or string)
Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getIntersection(anotherRange);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|anotherRange|Range o string|Objeto de intervalo o dirección de intervalo que se usará para determinar la intersección de los intervalos.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getIntersection("D4:G6");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!D4:F6
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getintersectionornullanotherrange-range-or-string"></a>getIntersectionOrNull(anotherRange: Range or string)
Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados. Si no se encuentra ninguna intersección, se devolverá un objeto NULL.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getIntersectionOrNull(anotherRange);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|anotherRange|Range o string|Objeto de intervalo o dirección de intervalo que se usará para determinar la intersección de los intervalos.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

### <a name="getlastcell"></a>getLastCell()
Obtiene la última celda del intervalo. Por ejemplo, la última celda de "B2:D5" es "D5".

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getLastCell();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getlastcolumn"></a>getLastColumn()
Obtiene la última columna del intervalo. Por ejemplo, la última columna de "B2:D5" es "D2:D5".

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getLastColumn();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastColumn();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F1:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getlastrow"></a>getLastRow()
Obtiene la última fila del intervalo. Por ejemplo, la última fila de "B2:D5" es "B5:D5".

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getLastRow();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastRow();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A8:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



### <a name="getoffsetrangerowoffset-number-columnoffset-number"></a>getOffsetRange(rowOffset: number, columnOffset: number)
Obtiene un objeto que representa un intervalo desplazado con respecto al intervalo especificado. La dimensión del intervalo devuelto coincidirá con este intervalo. Si el intervalo resultante se fuerza fuera de los límites de la cuadrícula de la hoja de cálculo, se producirá una excepción.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|rowOffset|number|Número de filas (número positivo, negativo o 0) que debe desplazarse el intervalo. Los valores positivos desplazan hacia abajo, mientras que los negativos lo hacen hacia arriba.|
|columnOffset|number|Número de columnas (número positivo, negativo o 0) que debe desplazarse el intervalo. Los valores positivos desplazan hacia la derecha, mientras que los negativos lo hacen hacia la izquierda.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:F6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getOffsetRange(-1,4);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!H3:K5
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getresizedrangedeltarows-number-deltacolumns-number"></a>getResizedRange(deltaRows: number, deltaColumns: number)
Obtiene un objeto Range similar al objeto Range actual, pero con su esquina inferior derecha expandida (o contraída) mediante un número de filas y columnas.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getResizedRange(deltaRows, deltaColumns);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|deltaRows|number|El número de filas en el que se va a expandir la esquina inferior derecha, con respecto al intervalo actual. Use un número positivo para expandir el intervalo, o un número negativo para reducirlo.|
|deltaColumns|number|El número de columnas en el que se va a expandir la esquina inferior derecha, con respecto al intervalo actual. Use un número positivo para expandir el intervalo, o un número negativo para reducirlo.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

### <a name="getrowrow-number"></a>getRow(row: number)
Obtiene una fila contenida en el intervalo.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getRow(row);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|row|number|Número de fila del intervalo que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A2:F2
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrowsabovecount-number"></a>getRowsAbove(count: number)
Obtiene un número determinado de filas encima del objeto Range actual.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getRowsAbove(count);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|count|number|Opcional. El número de filas que se va a incluir en el intervalo resultante. En general, use un número positivo para crear un intervalo fuera del intervalo actual. También puede usar un número negativo para crear un intervalo dentro del intervalo actual. El valor predeterminado es 1.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

### <a name="getrowsbelowcount-number"></a>getRowsBelow(count: number)
Obtiene un número determinado de filas debajo del objeto Range actual.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getRowsBelow(count);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|count|number|Opcional. El número de filas que se va a incluir en el intervalo resultante. En general, use un número positivo para crear un intervalo fuera del intervalo actual. También puede usar un número negativo para crear un intervalo dentro del intervalo actual. El valor predeterminado es 1.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

### <a name="getusedrangevaluesonly"></a>getUsedRange(valuesOnly)
Devuelve el intervalo usado del objeto de intervalo especificado.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getUsedRange(valuesOnly);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|valuesOnly|[ApiSet(Version|Solo tiene en cuenta las celdas con valores como celdas usadas.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeUR = range.getUsedRange();
    rangeUR.load('address');
    return ctx.sync().then(function() {
        console.log(rangeUR.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getvisibleview"></a>getVisibleView()
Representa las filas visibles del intervalo actual.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getVisibleView();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[RangeView](rangeview.md)

### <a name="insertshift-string"></a>insert(shift: string)
Inserta una celda o un intervalo de celdas en la hoja de cálculo en lugar de este intervalo y desplaza las demás celdas para crear espacio. Devuelve un objeto Range en el espacio que queda en blanco.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.insert(shift);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|Shift|string|Especifica hacia dónde se desplazarán las celdas.  Los valores posibles son: Down, Right|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js
    
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.insert();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


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

### <a name="mergeacross-bool"></a>merge(across: bool)
Combina las celdas del rango en una región de la hoja de cálculo.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.merge(across);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|across|bool|Opcional. Verdadero para que se combinen las celdas de cada fila del rango especificado como celdas combinadas distintas. El valor predeterminado es falso.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.merge(true);
    return ctx.sync(); 
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
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.unmerge();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="select"></a>select()
Selecciona el intervalo especificado en la interfaz de usuario de Excel.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.select();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.select();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="unmerge"></a>unmerge()
Separa las celdas del intervalo en celdas independientes.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.unmerge();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.unmerge();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

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
    var rangeName = 'MyRange';
    var range = ctx.workbook.names.getItem(rangeName).range;
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

En el ejemplo siguiente se establece el formato numérico, los valores y las fórmulas en una cuadrícula que contiene una cuadrícula de 2x3.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulas = [[null,null], [null,null], [null,"=G6-G5"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulas= formulas;
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Obtener la hoja de cálculo que contiene el intervalo. 

```js
/* This might be broken still - it was broken before because it 
    it was missing 'var', but might still be wrong because of
    getting information without loading properly. */
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    var range = namedItem.range;
    var rangeWorksheet = range.worksheet;
    rangeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(rangeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

