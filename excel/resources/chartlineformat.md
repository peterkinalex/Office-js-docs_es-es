# Objeto ChartLineFormat (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Office 2016_

Encapsula las opciones de formato para los elementos de línea.

## Propiedades

| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|color|string|Código de color HTML que representa el color de las líneas del gráfico.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Borra el formato de línea de un elemento de gráfico.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método

### clear()
Borra el formato de línea de un elemento de gráfico.

#### Sintaxis
```js
chartLineFormatObject.clear();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos

Borrar el formato de línea de las líneas de cuadrícula principales del eje de valores del gráfico denominado "Chart1".

```js
Excel.run(function (ctx) { 
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;	
	gridlines.format.line.clear();
	return ctx.sync().then(function() {
			console.log("Chart Major Gridlines Format Cleared");
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

Establecer las líneas de cuadrícula principales del eje de valores en rojo.

```js
Excel.run(function (ctx) { 
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.axes.valueaxis.majorGridlines;
	gridlines.format.line.color = "#FF0000";
	return ctx.sync().then(function() {
			console.log("Chart Gridlines Color Updated");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

