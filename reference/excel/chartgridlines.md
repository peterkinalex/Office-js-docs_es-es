# Objeto ChartGridlines (API de JavaScript para Excel)

Representa las líneas de división principales o secundarias del eje de un gráfico.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|visible|bool|Valor booleano que representa si las líneas de cuadrícula del eje están visibles o no.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|formato|[ChartGridlinesFormat](chartgridlinesformat.md)|Representa el formato de las líneas de cuadrícula del gráfico. Solo lectura.|

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

Obtener la propiedad `visible` de las líneas de cuadrícula principales del eje de valores de Chart1.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var majGridlines = chart.axes.valueaxis.majorGridlines;
    majGridlines.load('visible');
    return ctx.sync().then(function() {
            console.log(majGridlines.visible);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Establecer que se muestren las líneas de cuadrícula principales del eje de valores de Chart1.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.axes.valueaxis.majorgridlines.visible = true;
    return ctx.sync().then(function() {
            console.log("Axis Gridlines Added ");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
