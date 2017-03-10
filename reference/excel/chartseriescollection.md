# <a name="chartseriescollection-object-javascript-api-for-excel"></a>Objeto ChartSeriesCollection (API de JavaScript para Excel)

Representa una colección de series del gráfico.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|count|int|Devuelve el número de series incluidas en la colección. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[ChartSeries[]](chartseries.md)|Colección de objetos chartSeries. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|entero|Devuelve el número de series incluidas en la colección.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ChartSeries](chartseries.md)|Recupera una serie basada en su posición en la colección.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="getcount"></a>getCount()
Devuelve el número de series incluidas en la colección.

#### <a name="syntax"></a>Sintaxis
```js
chartSeriesCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
entero

### <a name="getitematindex-number"></a>getItemAt(index: number)
Recupera una serie basada en su posición en la colección.

#### <a name="syntax"></a>Sintaxis
```js
chartSeriesCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[ChartSeries](chartseries.md)

#### <a name="examples"></a>Ejemplos

Obtiene el nombre de la primera serie de la colección de series.

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('items');
    return ctx.sync().then(function() {
        console.log(seriesCollection.items[0].name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
Obtener los nombres de las series de la colección de series.

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < seriesCollection.items.length; i++)
        {
            console.log(seriesCollection.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Obtener el número de series del gráfico de la colección.

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('count');
    return ctx.sync().then(function() {
        console.log("series: Count= " + seriesCollection.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

