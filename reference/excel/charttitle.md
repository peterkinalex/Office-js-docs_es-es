# Objeto ChartTitle (API de JavaScript para Excel)

Representa un objeto de título de gráfico de un gráfico.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|overlay|bool|Valor booleano que representa si el título del gráfico se superpondrá al gráfico o no.|
|text|string|Representa el texto del título de un gráfico.|
|visible|bool|Valor booleano que representa la visibilidad de un objeto de título del gráfico.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|formato|[ChartTitleFormat](charttitleformat.md)|Representa el formato de un título del gráfico, que incluye el formato de relleno y de fuente. Solo lectura.|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
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
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### Valores devueltos
void
### Ejemplos de acceso a la propiedad

Obtener la propiedad `text` del título del gráfico de Chart1.

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 

var title = chart.title;
title.load('text');
return ctx.sync().then(function() {
        console.log(title.text);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Establecer la propiedad `text` del título del gráfico en "My Chart" y hacer que aparezca en la parte superior del gráfico sin que se superponga.

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 

chart.title.text= "My Chart"; 
chart.title.visible=true;
chart.title.overlay=true;

return ctx.sync().then(function() {
        console.log("Char Title Changed");
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
