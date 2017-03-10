# <a name="chartcollection-object-javascript-api-for-excel"></a>Objeto ChartCollection (API de JavaScript para Excel)

Colección de todos los objetos de gráfico en una hoja de cálculo.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|count|int|Devuelve el número de gráficos de la hoja de cálculo. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Chart[]](chart.md)|Colección de objetos de gráfico. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[add(type: string, sourceData: Range, seriesBy: string)](#addtype-string-sourcedata-range-seriesby-string)|[Chart](chart.md)|Crea un gráfico nuevo.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|entero|Devuelve el número de gráficos de la hoja de cálculo.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|Obtiene un gráfico mediante su nombre. Si hay varias tablas con el mismo nombre, se devolverá la primera.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|Obtiene un gráfico basado en su posición en la colección.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[Chart](chart.md)|Obtiene un gráfico mediante su nombre. Si hay varias tablas con el mismo nombre, se devolverá la primera.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="addtype-string-sourcedata-range-seriesby-string"></a>add(type: string, sourceData: Range, seriesBy: string)
Crea un nuevo gráfico.

#### <a name="syntax"></a>Sintaxis
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|type|string|Representa el tipo de un gráfico. Los valores posibles son: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|
|sourceData|Range|El objeto Range correspondiente a los datos de origen.|
|seriesBy|string|Opcional. Especifica la manera en que las columnas o las filas se usan como series de datos en el gráfico.  Los valores posibles son: Auto, Columns, Rows|

#### <a name="returns"></a>Valores devueltos
[Chart](chart.md)

#### <a name="examples"></a>Ejemplos

Agrega un gráfico de `chartType` "ColumnClustered" en la hoja de cálculo "Gráficos" con `sourceData` del rango "A1:B4" y `seriresBy` establecido en "auto".

```js
Excel.run(function (ctx) { 
    var rangeSelection = "A1:B4";
    var range = ctx.workbook.worksheets.getItem(sheetName)
        .getRange(rangeSelection);
    var chart = ctx.workbook.worksheets.getItem(sheetName)
        .charts.add("ColumnClustered", range, "auto");    return ctx.sync().then(function() {
            console.log("New Chart Added");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcount"></a>getCount()
Devuelve el número de gráficos de la hoja de cálculo.

#### <a name="syntax"></a>Sintaxis
```js
chartCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
entero

### <a name="getitemname-string"></a>getItem(name: string)
Obtiene un gráfico mediante su nombre. Si hay varias tablas con el mismo nombre, se devolverá la primera.

#### <a name="syntax"></a>Sintaxis
```js
chartCollectionObject.getItem(name);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Nombre del gráfico que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[Chart](chart.md)

#### <a name="examples"></a>Ejemplos

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


#### <a name="examples"></a>Ejemplos

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



#### <a name="examples"></a>Ejemplos

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


### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtiene un gráfico basado en su posición en la colección.

#### <a name="syntax"></a>Sintaxis
```js
chartCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[Chart](chart.md)

#### <a name="examples"></a>Ejemplos

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


### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
Obtiene un gráfico mediante su nombre. Si hay varias tablas con el mismo nombre, se devolverá la primera.

#### <a name="syntax"></a>Sintaxis
```js
chartCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Nombre del gráfico que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[Chart](chart.md)
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < charts.items.length; i++)
        {
            console.log(charts.items[i].name);
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

