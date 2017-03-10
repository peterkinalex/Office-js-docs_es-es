# <a name="chartpointscollection-object-javascript-api-for-excel"></a>Objeto ChartPointsCollection (API de JavaScript para Excel)

Colección de todos los puntos del gráfico dentro de una serie de un gráfico.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|count|entero|Devuelve el número de puntos del gráfico de la serie. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[ChartPoint[]](chartpoint.md)|Colección de objetos chartPoint. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|entero|Devuelve el número de puntos del gráfico de la serie.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ChartPoint](chartpoint.md)|Recupera un punto basándose en su posición dentro de la serie.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="getcount"></a>getCount()
Devuelve el número de puntos del gráfico de la serie.

#### <a name="syntax"></a>Sintaxis
```js
chartPointsCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
entero

### <a name="getitematindex-number"></a>getItemAt(index: number)
Recupera un punto basándose en su posición dentro de la serie.

#### <a name="syntax"></a>Sintaxis
```js
chartPointsCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[ChartPoint](chartpoint.md)

#### <a name="examples"></a>Ejemplos
Establece el color de borde de los primeros puntos de la colección de puntos.

```js
Excel.run(function (ctx) { 
    var points = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
    points.getItemAt(0).format.fill.setSolidColor("8FBC8F");
    return ctx.sync().then(function() {
        console.log("Point Border Color Changed");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```### Property access examples

Get the names of points in the points collection

```js
Excel.run(function (ctx) { 
    var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
    pointsCollection.load('items');
    return ctx.sync().then(function() {
        console.log("Points Collection loaded");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Obtiene el número de puntos.

```js
Excel.run(function (ctx) { 
    var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
    pointsCollection.load('count');
    return ctx.sync().then(function() {
        console.log("points: Count= " + pointsCollection.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
