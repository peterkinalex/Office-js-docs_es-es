# Objeto Range (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Office 2016_

El intervalo representa un conjunto de una o más celdas contiguas, como una celda, una fila, una columna, un bloque de celdas, etc.

## Propiedades

| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|address|string|Representa la referencia de intervalo en estilo A1. El valor de dirección contendrá la referencia de hoja (por ejemplo, Sheet1!A1:B4). Solo lectura.|
|addressLocal|string|Representa la referencia del intervalo especificado en el idioma del usuario. Solo lectura.|
|cellCount|int|Número de celdas del intervalo. Solo lectura.|
|columnCount|int|Representa el número total de columnas del intervalo. Solo lectura.|
|columnIndex|int|Representa el número de columna de la primera celda del intervalo. Indizado con cero. Solo lectura.|
|formulas|object[][]|Representa la fórmula en notación de estilo A1.|
|formulasLocal|object[][]|Representa la fórmula en notación de estilo A1, en el idioma del usuario y en la configuración regional del formato numérico. Por ejemplo, la fórmula "=SUM(A1, 1.5)" en inglés se convertiría en "=SUMME(A1; 1,5)" en alemán.|
|numberFormat|object[][]|Representa el código de formato numérico para la celda especificada.|
|rowCount|int|Devuelve el número total de filas del intervalo. Solo lectura.|
|rowIndex|int|Devuelve el número de fila de la primera celda del intervalo. Indizado con cero. Solo lectura.|
|text|object[][]|Valores de texto del intervalo especificado. El valor Text no dependerá del ancho de la celda. La sustitución del signo # que tiene lugar en la interfaz de usuario de Excel no afectará al valor de texto devuelto por la API. Solo lectura.|
|valueTypes|string|Representa el tipo de datos de cada celda. Solo lectura. Los valores posibles son: Unknown, Empty, String, Integer, Double, Boolean, Error.|
|values|object[][]|Representa los valores sin formato del intervalo especificado. Los datos devueltos pueden ser de tipo cadena, número o booleano. La celda que contenga un error devolverá la cadena de error.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
| Relación | Tipo|Descripción|
|:---------------|:--------|:----------|
|formato|[RangeFormat](rangeformat.md)|Devuelve un objeto de formato que encapsula la fuente, el relleno, los bordes, la alineación y otras propiedades del intervalo. Solo lectura.|
|worksheet|[Worksheet](worksheet.md)|Hoja de cálculo que contiene el intervalo actual. Solo lectura.|

## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[clear(applyTo: string)](#clearapplyto-string)|void|Borra valores de intervalo, formato, relleno, borde, etc.|
|[delete(shift: string)](#deleteshift-string)|void|Elimina las celdas asociadas al intervalo.|
|[getBoundingRect(anotherRange: Range or string)](#getboundingrectanotherrange-range-or-string)|[Range](range.md)|Obtiene el objeto de intervalo más pequeño que abarca los intervalos especificados. Por ejemplo, el valor getBoundingRect de "B2:C5" y "D10:E15" es "B2:E15".|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Obtiene el objeto de intervalo que contiene la celda en función de los números de fila y columna. La celda puede estar fuera de los límites del intervalo principal, siempre y cuando permanezca dentro de la cuadrícula de la hoja de cálculo. La celda devuelta se ubica con respecto a la celda superior izquierda del intervalo.|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|Obtiene una columna contenida en el intervalo.|
|[getEntireColumn()](#getentirecolumn)|[Range](range.md)|Obtiene un objeto que representa toda la columna del intervalo.|
|[getEntireRow()](#getentirerow)|[Range](range.md)|Obtiene un objeto que representa toda la fila del intervalo.|
|[getIntersection(anotherRange: Range or string)](#getintersectionanotherrange-range-or-string)|[Range](range.md)|Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados.|
|[getLastCell()](#getlastcell)|[Range](range.md)|Obtiene la última celda del intervalo. Por ejemplo, la última celda de "B2:D5" es "D5".|
|[getLastColumn()](#getlastcolumn)|[Range](range.md)|Obtiene la última columna del intervalo. Por ejemplo, la última columna de "B2:D5" es "D2:D5".|
|[getLastRow()](#getlastrow)|[Range](range.md)|Obtiene la última fila del intervalo. Por ejemplo, la última fila de "B2:D5" es "B5:D5".|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|Obtiene un objeto que representa un intervalo desplazado con respecto al intervalo especificado. La dimensión del intervalo devuelto coincidirá con este intervalo. Si el intervalo resultante se fuerza fuera de los límites de la cuadrícula de la hoja de cálculo, se producirá una excepción.|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|Obtiene una fila contenida en el intervalo.|
|[getUsedRange()](#getusedrange)|[Range](range.md)|Devuelve el intervalo usado del objeto de intervalo especificado.|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|Inserta una celda o un intervalo de celdas en la hoja de cálculo en lugar de este intervalo y desplaza las demás celdas para crear espacio. Devuelve un objeto Range en el espacio que queda en blanco.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|
|[select()](#select)|void|Selecciona el intervalo especificado en la interfaz de usuario de Excel.|

## Detalles del método

### clear(applyTo: string)
Borra valores de intervalo, formato, relleno, borde, etc.

#### Sintaxis
```js
rangeObject.clear(applyTo);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|applyTo|string|Opcional. Determina el tipo de acción de borrado. Los valores posibles son: `All` (opción predeterminada), `Formats`, `Contents`|

#### Valores devueltos
void

#### Ejemplos

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

### delete(shift: string)
Elimina las celdas asociadas al intervalo.

#### Sintaxis
```js
rangeObject.delete(shift);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|Shift|string|Especifica hacia dónde se desplazarán las celdas.  Los valores posibles son: Up, Left|

#### Valores devueltos
void

#### Ejemplos

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

### getBoundingRect(anotherRange: Range or string)
Obtiene el objeto de intervalo más pequeño que abarca los intervalos especificados. Por ejemplo, el valor GetBoundingRect de "B2:C5" y "D10:E15" es "B2:E15".

#### Sintaxis
```js
rangeObject.getBoundingRect(anotherRange);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|anotherRange|Range o string|Objeto o dirección de intervalo o nombre de intervalo.|

#### Valores devueltos
[Range](range.md)

#### Ejemplos

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

### getCell(row: number, column: number)
Obtiene el objeto de intervalo que contiene la celda en función de los números de fila y columna. La celda puede estar fuera de los límites del intervalo principal, siempre y cuando permanezca dentro de la cuadrícula de la hoja de cálculo. La celda devuelta se ubica con respecto a la celda superior izquierda del intervalo.

#### Sintaxis
```js
rangeObject.getCell(row, column);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
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

### getColumn(column: number)
Obtiene una columna contenida en el intervalo.

#### Sintaxis
```js
rangeObject.getColumn(column);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|column|number|Número de columna del intervalo que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[Range](range.md)

#### Ejemplos

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

### getEntireColumn()
Obtiene un objeto que representa toda la columna del intervalo.

#### Sintaxis
```js
rangeObject.getEntireColumn();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos

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
### getEntireRow()
Obtiene un objeto que representa toda la fila del intervalo.

#### Sintaxis
```js
rangeObject.getEntireRow();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos
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

### getIntersection(anotherRange: Range or string)
Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados.

#### Sintaxis
```js
rangeObject.getIntersection(anotherRange);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|anotherRange|Range o string|Objeto de intervalo o dirección de intervalo que se usará para determinar la intersección de los intervalos.|

#### Valores devueltos
[Range](range.md)

#### Ejemplos

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

### getLastCell()
Obtiene la última celda del intervalo. Por ejemplo, la última celda de "B2:D5" es "D5".

#### Sintaxis
```js
rangeObject.getLastCell();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos

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

### getLastColumn()
Obtiene la última columna del intervalo. Por ejemplo, la última columna de "B2:D5" es "D2:D5".

#### Sintaxis
```js
rangeObject.getLastColumn();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos

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

### getLastRow()
Obtiene la última fila del intervalo. Por ejemplo, la última fila de "B2:D5" es "B5:D5".

#### Sintaxis
```js
rangeObject.getLastRow();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos

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


### getOffsetRange(rowOffset: number, columnOffset: number)
Obtiene un objeto que representa un intervalo desplazado con respecto al intervalo especificado. La dimensión del intervalo devuelto coincidirá con este intervalo. Si el intervalo resultante se fuerza fuera de los límites de la cuadrícula de la hoja de cálculo, se producirá una excepción.

#### Sintaxis
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|rowOffset|number|Número de filas (número positivo, negativo o 0) que debe desplazarse el intervalo. Los valores positivos desplazan hacia abajo, mientras que los negativos lo hacen hacia arriba.|
|columnOffset|number|Número de columnas (número positivo, negativo o 0) que debe desplazarse el intervalo. Los valores positivos desplazan hacia la derecha, mientras que los negativos lo hacen hacia la izquierda.|

#### Valores devueltos
[Range](range.md)

#### Ejemplos

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

### getRow(row: number)
Obtiene una fila contenida en el intervalo.

#### Sintaxis
```js
rangeObject.getRow(row);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|row|number|Número de fila del intervalo que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[Range](range.md)

#### Ejemplos

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

### getUsedRange()
Devuelve el intervalo usado del objeto de intervalo especificado.

#### Sintaxis
```js
rangeObject.getUsedRange();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos

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

### insert(shift: string)
Inserta una celda o un intervalo de celdas en la hoja de cálculo en lugar de este intervalo y desplaza las demás celdas para crear espacio. Devuelve un objeto Range en el espacio que queda en blanco.

#### Sintaxis
```js
rangeObject.insert(shift);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|Shift|string|Especifica hacia dónde se desplazarán las celdas.  Los valores posibles son: Down, Right|

#### Valores devueltos
[Range](range.md)

#### Ejemplos

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

### load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### Sintaxis
```js
object.load(param);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### Valores devueltos
void
### select()
Selecciona el intervalo especificado en la interfaz de usuario de Excel.

#### Sintaxis
```js
rangeObject.select();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos

```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "F5:F10"; 
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.select();
	return ctx.sync(); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Ejemplos de acceso a la propiedad

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
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var namedItem = names.getItem('MyRange');
	range = namedItem.range;
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

