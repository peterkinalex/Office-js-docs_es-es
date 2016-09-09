# Objeto ChartFill (API de JavaScript para Excel)

Representa el formato de relleno para un elemento de gráfico.

## Propiedades

Ninguno

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Borra el color de relleno de un elemento de gráfico.|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|Establece el formato de relleno de un elemento de gráfico en un color uniforme.|

## Detalles del método


### clear()
Borra el color de relleno de un elemento de gráfico.

#### Sintaxis
```js
chartFillObject.clear();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos

Borrar el formato de línea de las líneas de cuadrícula principales del eje de valores del gráfico denominado "Chart1".

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;   
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

### setSolidColor(color: string)
Establece el formato de relleno de un elemento de gráfico en un color uniforme.

#### Sintaxis
```js
chartFillObject.setSolidColor(color);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|color|string|Código de color HTML que representa el color de la línea de borde con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|

#### Valores devueltos
void

#### Ejemplos

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
