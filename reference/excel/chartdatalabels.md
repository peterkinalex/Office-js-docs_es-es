# <a name="chartdatalabels-object-(javascript-api-for-excel)"></a>Objeto ChartDataLabels (API de JavaScript para Excel)

Representa una colección de todas las etiquetas de datos en un punto del gráfico.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|position|string|Valor DataLabelPosition que representa la posición de la etiqueta de datos. Los valores posibles son: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout. Solo escritura.|
|Separator|string|Cadena que representa el separador empleado para las etiquetas de datos de un gráfico. Solo escritura.|
|showBubbleSize|bool|Valor booleano que representa si el tamaño de la burbuja de la etiqueta de datos es visible o no. Solo escritura.|
|showCategoryName|bool|Valor booleano que representa si el nombre de categoría de la etiqueta de datos es visible o no. Solo escritura.|
|showLegendKey|bool|Valor booleano que representa si la clave de leyenda de la etiqueta de datos es visible o no. Solo escritura.|
|showPercentage|bool|Valor booleano que representa si el porcentaje de la etiqueta de datos es visible o no. Solo escritura.|
|showSeriesName|bool|Valor booleano que representa si el nombre de serie de la etiqueta de datos es visible o no. Solo escritura.|
|showValue|bool|Valor booleano que representa si el valor de la etiqueta de datos es visible o no. Solo escritura.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|Representa el formato de las etiquetas de datos del gráfico, que incluye el formato de relleno y de fuente. Solo lectura.|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## <a name="method-details"></a>Detalles del método


### <a name="load(param:-object)"></a>load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
object.load(param);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
void
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

Hacer que el nombre de serie se muestre en las etiquetas de datos y establecer la propiedad `position` de las etiquetas de datos en "top".

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.datalabels.visible = true;
    chart.datalabels.position = "top";
    chart.datalabels.ShowSeriesName = true;
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
