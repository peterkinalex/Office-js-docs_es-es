# <a name="charttitle-object-javascript-api-for-excel"></a>Objeto ChartTitle (API de JavaScript para Excel)

Representa un objeto de título de gráfico de un gráfico.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|overlay|bool|Valor booleano que representa si el título del gráfico se superpondrá al gráfico o no.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|text|string|Representa el texto del título de un gráfico.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Un valor booleano que representa la visibilidad de un objeto de título del gráfico.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|format|[ChartTitleFormat](charttitleformat.md)|Representa el formato de un título del gráfico, que incluye el formato de relleno y de fuente. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos
Ninguno


## <a name="method-details"></a>Detalles del método

### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

Obtener la propiedad `text` del título del gráfico de Gráfico1.

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
});
```

Establecer la propiedad `text` del título del gráfico en "Mi gráfico" y hacer que aparezca en la parte superior del gráfico sin que se superponga.

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
});
```
