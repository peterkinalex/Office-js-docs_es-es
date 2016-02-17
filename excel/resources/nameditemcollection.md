# Objeto NamedItemCollection (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Office 2016_

Colección de todos los objetos namedItem que forman parte del libro.

## Propiedades

| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|Items|[NamedItem[]](nameditem.md)|Colección de objetos namedItem. Solo lectura.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Obtiene un objeto namedItem mediante su nombre.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método

### getItem(name: string)
Obtiene un objeto namedItem mediante su nombre.

#### Sintaxis
```js
namedItemCollectionObject.getItem(name);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|name|string|Nombre de namedItem.|

#### Valores devueltos
[NamedItem](nameditem.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
	var nameditem = ctx.workbook.names.getItem(wSheetName);
	nameditem.load('type');
	return ctx.sync().then(function() {
			console.log(nameditem.type);
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
	var nameditems = ctx.workbook.names;
	nameditems.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < nameditems.items.length; i++)
		{
			console.log(nameditems.items[i].name);
			console.log(nameditems.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Obtener el número de objetos namedItem.

```js
Excel.run(function (ctx) { 
	var nameditems = ctx.workbook.names;
	nameditems.load('count');
	return ctx.sync().then(function() {
		console.log("nameditems: Count= " + nameditems.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


