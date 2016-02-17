# Objeto Binding (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Office 2016_

Representa un enlace de Office.js que se define en el libro.

## Propiedades

| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|id|string|Representa el identificador de enlace. Solo lectura.|
|type|string|Devuelve el tipo de enlace. Solo lectura. Los valores posibles son: Range, Table, Text.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|Devuelve el intervalo representado por el enlace. Se producirá un error si el enlace no es del tipo correcto.|
|[getTable()](#gettable)|[Table](table.md)|Devuelve la tabla representada por el enlace. Se producirá un error si el enlace no es del tipo correcto.|
|[getText()](#gettext)|string|Devuelve el texto representado por el enlace. Se producirá un error si el enlace no es del tipo correcto.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método

### getRange()
Devuelve el intervalo representado por el enlace. Se producirá un error si el enlace no es del tipo correcto.

#### Sintaxis
```js
bindingObject.getRange();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos
El ejemplo siguiente usa el objeto de enlace para obtener el intervalo asociado.

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var range = binding.getRange();
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

### getTable()
Devuelve la tabla representada por el enlace. Se producirá un error si el enlace no es del tipo correcto.

#### Sintaxis
```js
bindingObject.getTable();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Table](table.md)

#### Ejemplos
```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var table = binding.getTable();
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

### getText()
Devuelve el texto representado por el enlace. Se producirá un error si el enlace no es del tipo correcto.

#### Sintaxis
```js
bindingObject.getText();
```

#### Parámetros
Ninguno

#### Valores devueltos
string

#### Ejemplos

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var text = binding.getText();
	ctx.load('text');
	return ctx.sync().then(function() {
		console.log(text);
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
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, acepta un objeto [loadOption](loadoption.md).|

#### Valores devueltos
void
### Ejemplos de acceso a la propiedad

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	binding.load('type');
	return ctx.sync().then(function() {
		console.log(binding.type);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

