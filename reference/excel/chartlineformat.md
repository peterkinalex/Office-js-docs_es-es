# <a name="chartlineformat-object-javascript-api-for-excel"></a>Objeto ChartLineFormat (API de JavaScript para Excel)

Encapsula las opciones de formato para los elementos de línea.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|color|string|Código de color HTML que representa el color de las líneas del gráfico.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Borra el formato de línea de un elemento de gráfico.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="clear"></a>clear()
Borra el formato de línea de un elemento de gráfico.

#### <a name="syntax"></a>Sintaxis
```js
chartLineFormatObject.clear();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos

Borra el formato de línea de las líneas de cuadrícula principales del eje de valores del gráfico denominado "Chart1".

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;    
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
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

Establecer las líneas de cuadrícula principales del eje de valores en rojo.

```js
Excel.run(function (ctx) {
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;
    gridlines.format.line.color = "#FF0000";
    return ctx.sync().then(function () {
        console.log("Chart Gridlines Color Updated");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
