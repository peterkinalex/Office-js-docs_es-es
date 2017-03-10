# <a name="chartfill-object-javascript-api-for-excel"></a>Objeto ChartFill (API de JavaScript para Excel)

Representa el formato de relleno para un elemento de gráfico.

## <a name="properties"></a>Propiedades

Ninguno

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Borra el color de relleno de un elemento de gráfico.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|Establece el formato de relleno de un elemento de gráfico en un color uniforme.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="clear"></a>clear()
Borra el color de relleno de un elemento de gráfico.

#### <a name="syntax"></a>Sintaxis
```js
chartFillObject.clear();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos

Borrar el formato de línea de las líneas de cuadrícula principales del eje de valores del gráfico denominado "Chart1".

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

### <a name="setsolidcolorcolor-string"></a>setSolidColor(color: string)
Establece el formato de relleno de un elemento de gráfico en un color uniforme.

#### <a name="syntax"></a>Sintaxis
```js
chartFillObject.setSolidColor(color);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|color|string|Código de color HTML que representa el color de la línea de borde con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos

Establecer el color de fondo de Chart1 en rojo.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    

    chart.format.fill.setSolidColor("#FF0000");

    return ctx.sync().then(function() {
            console.log("Chart1 Background Color Changed.");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
