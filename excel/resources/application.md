# Objeto Application (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Office 2016_

Representa la aplicación de Excel que administra el libro.

## Propiedades

| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|calculationMode|string|Devuelve el modo de cálculo usado en el libro. Solo lectura. Los valores posibles son: `Automatic` Excel controla el recálculo; `AutomaticExceptTables` Excel controla el recálculo pero omite los cambios de las tablas; `Manual` el cálculo se realiza cuando el usuario lo solicita.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Recalcula todos los libros abiertos actualmente en Excel.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método

### calculate(calculationType: string)
Recalcula todos los libros abiertos actualmente en Excel.

#### Sintaxis
```js
applicationObject.calculate(calculationType);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|calculationType|string|Especifica el tipo de cálculo que se va a usar. Los valores posibles son: `Recalculate` opción predeterminada, realiza el cálculo normal calculando todas las fórmulas del libro; `Full` fuerza el cálculo completo de los datos; `FullRebuild` fuerza el cálculo completo de los datos y vuelve a crear las dependencias.|

#### Valores devueltos
void

#### Ejemplos
```js
Excel.run(function (ctx) { 
	ctx.workbook.application.calculate('Full');
	return ctx.sync(); 
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
	var application = ctx.workbook.application;
	application.load('calculationMode');
	return ctx.sync().then(function() {
		console.log(application.calculationMode);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


