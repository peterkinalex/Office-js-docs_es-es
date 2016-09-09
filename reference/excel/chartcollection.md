# Objeto ChartCollection (API de JavaScript para Excel)

Colección de todos los objetos de gráfico en una hoja de cálculo.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|count|int|Devuelve el número de gráficos de la hoja de cálculo. Solo lectura.|
|Items|[Chart[]](chart.md)|Colección de objetos de gráfico. Solo lectura.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[add(type: string, sourceData: Range, seriesBy: string)](#addtype-string-sourcedata-range-seriesby-string)|[Chart](chart.md)|Crea un nuevo gráfico.|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|Obtiene un gráfico mediante su nombre. Si hay varios gráficos con el mismo nombre, se devuelve el primero.|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|Obtiene un gráfico basado en su posición en la colección.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método


### add(type: string, sourceData: Range, seriesBy: string)
Crea un nuevo gráfico.

#### Sintaxis
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|type|string|Representa el tipo de un gráfico. Los valores posibles son: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|
|sourceData|Range|Objeto de intervalo que contiene los datos de origen.|
|seriesBy|string|Opcional. Especifica la manera en que las columnas o las filas se usan como series de datos en el gráfico.  Los valores posibles son: Auto, Columns, Rows|

#### Valores devueltos
[Chart](chart.md)

#### Ejemplos

Agregar un gráfico de `chartType` "ColumnClustered" en la hoja de cálculo "Gráficos" con `sourceData` del intervalo "A1:B4" y `seriesBy` establecido en "auto".

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var sourceData = sheetName + "!" + "A1:B4";
    var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("ColumnClustered", sourceData, "auto");
    return ctx.sync().then(function() {
            console.log("New Chart Added");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItem(name: string)
Obtiene un gráfico mediante su nombre. Si hay varios gráficos con el mismo nombre, se devuelve el primero.

#### Sintaxis
```js
chartCollectionObject.getItem(name);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|name|string|Nombre del gráfico que se va a recuperar.|

#### Valores devueltos
[Chart](chart.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
    var chartname = 'Chart1';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartname);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


#### Ejemplos

```js
Excel.run(function (ctx) { 
    var chartId = 'SamplChartId';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartId);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



#### Ejemplos

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItemAt(index: number)
Obtiene un gráfico basado en su posición en la colección.

#### Sintaxis
```js
chartCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[Chart](chart.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


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

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < charts.items.length; i++)
        {
            console.log(charts.items[i].name);
            console.log(charts.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Obtener el número de gráficos.

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('count');
    return ctx.sync().then(function() {
        console.log("charts: Count= " + charts.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

