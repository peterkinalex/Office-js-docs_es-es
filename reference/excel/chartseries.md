# <a name="chartseries-object-javascript-api-for-excel"></a>Objeto ChartSeries (API de JavaScript para Excel)

Representa una serie de un gráfico.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|name|string|Representa el nombre de una serie de un gráfico.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|format|[ChartSeriesFormat](chartseriesformat.md)|Representa el formato de una serie del gráfico, que incluye el formato de relleno y de línea. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|points|[ChartPointsCollection](chartpointscollection.md)|Representa una colección de todos los puntos de la serie. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos
Ninguno


## <a name="method-details"></a>Detalles del método

### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

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
