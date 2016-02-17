# Objeto ChartSeries (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Office 2016_

Representa una serie de un gráfico.

## Propiedades

| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|name|string|Representa el nombre de una serie de un gráfico.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
| Relación | Tipo|Descripción|
|:---------------|:--------|:----------|
|formato|[ChartSeriesFormat](chartseriesformat.md)|Representa el formato de una serie del gráfico, que incluye el formato de relleno y de línea. Solo lectura.|
|Puntos|[ChartPointsCollection](chartpointscollection.md)|Representa una colección de todos los puntos de la serie. Solo lectura.|

## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método

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

Cambiar el nombre de la primera serie de Chart1 a "New Series Name".

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.series.getItemAt(0).name = "New Series Name";
	return ctx.sync().then(function() {
			console.log("Series1 Renamed");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

