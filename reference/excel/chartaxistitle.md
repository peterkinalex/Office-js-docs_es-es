# <a name="chartaxistitle-object-javascript-api-for-excel"></a>Objeto ChartAxisTitle (API de JavaScript para Excel)

Representa el título del eje de un gráfico.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|text|string|Representa el título del eje.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Un valor booleano que especifica la visibilidad del título de un eje.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|format|[ChartAxisTitleFormat](chartaxistitleformat.md)|Representa el formato del título del eje del gráfico. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos
Ninguno


## <a name="method-details"></a>Detalles del método

### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
Obtener la propiedad `text` del título del eje de gráfico a partir del eje de valores de Gráfico1.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    var title = chart.axes.valueAxis.title;
    title.load('text');
    return ctx.sync().then(function() {
            console.log(title.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Agregar "Values" como título del eje de valores.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.axes.valueAxis.title.text = "Values";
    return ctx.sync().then(function() {
            console.log("Axis Title Added ");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
