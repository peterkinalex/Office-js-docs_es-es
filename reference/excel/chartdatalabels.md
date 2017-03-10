# <a name="chartdatalabels-object-javascript-api-for-excel"></a>Objeto ChartDataLabels (API de JavaScript para Excel)

Representa una colección de todas las etiquetas de datos en un punto del gráfico.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|position|string|Valor DataLabelPosition que representa la posición de la etiqueta de datos. Los valores posibles son: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|separator|string|Cadena que representa el separador empleado para las etiquetas de datos de un gráfico.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBubbleSize|bool|Valor booleano que representa si el tamaño de la burbuja de la etiqueta de datos es visible o no.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showCategoryName|bool|Valor booleano que representa si el nombre de categoría de la etiqueta de datos es visible o no.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showLegendKey|bool|Valor booleano que representa si la clave de leyenda de la etiqueta de datos es visible o no.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showPercentage|bool|Valor booleano que representa si el porcentaje de la etiqueta de datos es visible o no.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showSeriesName|bool|Valor booleano que representa si el nombre de serie de la etiqueta de datos es visible o no.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showValue|bool|Valor booleano que representa si el valor de la etiqueta de datos es visible o no.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|Representa el formato de las etiquetas de datos del gráfico, que incluye el formato de relleno y de fuente. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos
Ninguno


## <a name="method-details"></a>Detalles del método

### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

Hacer que el nombre de serie se muestre en DataLabels y establecer la propiedad `position` de DataLabels en "top".

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.datalabels.showValue = true;
    chart.datalabels.position = "top";
    chart.datalabels.showSeriesName = true;
    return ctx.sync().then(function() {
            console.log("Datalabels Shown");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
