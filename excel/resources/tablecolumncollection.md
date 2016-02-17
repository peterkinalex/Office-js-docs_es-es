# Objeto TableColumnCollection (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Office 2016_

Representa una colección de todas las columnas que forman parte de la tabla.

## Propiedades

| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|count|int|Devuelve el número de columnas de la tabla. Solo lectura.|
|Items|[TableColumn[]](tablecolumn.md)|Colección de objetos tableColumn. Solo lectura.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableColumn](tablecolumn.md)|Agrega una nueva columna a la tabla.|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|Obtiene un objeto de columna por nombre o identificador.|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|Obtiene una columna basada en su posición en la colección.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método

### add(index: number, values: (boolean or string or number)[][])
Agrega una nueva columna a la tabla.

#### Sintaxis
```js
tableColumnCollectionObject.add(index, values);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|index|number|Especifica la posición relativa de la nueva columna. La columna anterior en esta posición se desplaza hacia la derecha. El valor de índice debe ser igual o menor que el valor de índice de la última columna, por lo que no puede usarse para agregar una columna al final de la tabla. Indexado con cero.|
|values|(boolean or string or number)[][]|Opcional. Matriz bidimensional de valores sin formato de la columna de la tabla.|

#### Valores devueltos
[TableColumn](tablecolumn.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
	var column = tables.getItem("Table1").columns.add(null, values);
	column.load('name');
	return ctx.sync().then(function() {
		console.log(column.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItem(key: number or string)
Obtiene un objeto de columna por nombre o identificador.

#### Sintaxis
```js
tableColumnCollectionObject.getItem(key);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|Key|número o cadena| Nombre o identificador de columna.|

#### Valores devueltos
[TableColumn](tablecolumn.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
	var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItem(0);
	tablecolumn.load('name');
	return ctx.sync().then(function() {
			console.log(tablecolumn.name);
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
	var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
	tablecolumn.load('name');
	return ctx.sync().then(function() {
			console.log(tablecolumn.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getItemAt(index: number)
Obtiene una columna basada en su posición en la colección.

#### Sintaxis
```js
tableColumnCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[TableColumn](tablecolumn.md)

#### Ejemplos
```js
Excel.run(function (ctx) { 
	var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
	tablecolumn.load('name');
	return ctx.sync().then(function() {
			console.log(tablecolumn.name);
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
### Ejemplos de acceso a la propiedad

```js
Excel.run(function (ctx) { 
	var tablecolumns = ctx.workbook.tables.getItem['Table1'].columns;
	tablecolumns.load('items');
	return ctx.sync().then(function() {
		console.log("tablecolumns Count: " + tablecolumns.count);
		for (var i = 0; i < tablecolumns.items.length; i++)
		{
			console.log(tablecolumns.items[i].name);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
