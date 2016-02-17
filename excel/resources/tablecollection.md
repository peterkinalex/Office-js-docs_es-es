# Objeto TableCollection (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Office 2016_

Representa una colección de todas las tablas que forman parte del libro.

## Propiedades

| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|count|int|Devuelve el número de tablas del libro. Solo lectura.|
|Items|[Table[]](table.md)|Colección de objetos de tabla. Solo lectura.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[add(address: string, hasHeaders: bool)](#addaddress-string-hasheaders-bool)|[Table](table.md)|Crea una tabla nueva. La dirección de origen del intervalo determina la hoja de cálculo en la que se agregará la tabla. Si no se puede agregar la tabla (por ejemplo, porque la dirección no es válida o porque la tabla se superpondría con otra tabla), se produce un error.|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Table](table.md)|Obtener una tabla por nombre o identificador.|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|Obtiene una tabla basada en su posición en la colección.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método

### add(address: string, hasHeaders: bool)
Crea una tabla nueva. La dirección de origen del intervalo determina la hoja de cálculo en la que se agregará la tabla. Si no se puede agregar la tabla (por ejemplo, porque la dirección no es válida o porque la tabla se superpondría con otra tabla), se produce un error.

#### Sintaxis
```js
tableCollectionObject.add(address, hasHeaders);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|address|string|Dirección o nombre del objeto de intervalo que representa el origen de datos. Si la dirección no contiene un nombre de hoja, se usa la hoja activa en ese momento.|
|hasHeaders|bool|Valor booleano que indica si los datos que se están importando tienen etiquetas de columna. Si el origen no contiene encabezados (es decir, cuando esta propiedad se establece en false), Excel generará automáticamente un encabezado desplazando los datos hacia abajo una fila.|

#### Valores devueltos
[Table](table.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
	table.load('name');
	return ctx.sync().then(function() {
		console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getItem(key: number or string)
Obtener una tabla por nombre o identificador.

#### Sintaxis
```js
tableCollectionObject.getItem(key);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|Key|number o string|Nombre o identificador de la tabla que se va a recuperar.|

#### Valores devueltos
[Table](table.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	return ctx.sync().then(function() {
			console.log(table.index);
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
	var table = ctx.workbook.tables.getItemAt(0);
	return ctx.sync().then(function() {
			console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItemAt(index: number)
Obtiene una tabla basada en su posición en la colección.

#### Sintaxis
```js
tableCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[Table](table.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItemAt(0);
	return ctx.sync().then(function() {
			console.log(table.name);
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
	var tables = ctx.workbook.tables;
	tables.load('items');
	return ctx.sync().then(function() {
		console.log("tables Count: " + tables.count);
		for (var i = 0; i < tables.items.length; i++)
		{
			console.log(tables.items[i].name);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Obtener el número de tablas.

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	tables.load('count');
	return ctx.sync().then(function() {
		console.log(tables.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
