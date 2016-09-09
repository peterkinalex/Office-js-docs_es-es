# Objeto ChartLegend (API de JavaScript para Excel)

Representa la leyenda de un gráfico.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|overlay|bool|Valor booleano que indica si la leyenda del gráfico debe superponerse al cuerpo principal del gráfico.|
|position|string|Representa la posición de la leyenda del gráfico. Los valores posibles son: Top, Bottom, Left, Right, Corner, Custom.|
|visible|bool|Valor booleano que representa la visibilidad de un objeto ChartLegend.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|formato|[ChartLegendFormat](chartlegendformat.md)|Representa el formato de una leyenda del gráfico, que incluye el formato de relleno y de fuente. Solo lectura.|

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

Obtener la propiedad `position` de la leyenda del gráfico de Chart1.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var legend = chart.legend;
    legend.load('position');
    return ctx.sync().then(function() {
            console.log(legend.position);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Establecer esta opción para mostrar la leyenda de Chart1 en la parte superior del gráfico.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.legend.visible = true;
    chart.legend.position = "top"; 
    chart.legend.overlay = false; 
    return ctx.sync().then(function() {
            console.log("Legend Shown ");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
``` 
