# <a name="tablerowcollection-object-javascript-api-for-excel"></a>Objeto TableRowCollection (API de JavaScript para Excel)

Representa una colección de todas las filas que forman parte de la tabla.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|count|int|Devuelve el número de filas de la tabla. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[TableRow[]](tablerow.md)|Colección de objetos tableRow. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableRow](tablerow.md)|Agrega una o más filas a la tabla. El objeto devuelto será el superior de las filas recién agregadas.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|entero|Obtiene el número de filas de la tabla.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|Obtiene una fila basada en su posición en la colección.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="addindex-number-values-boolean-or-string-or-number"></a>add(index: number, values: (boolean or string or number)[][])
Agrega una o más filas a la tabla. El objeto devuelto será el superior de las filas recién agregadas.

#### <a name="syntax"></a>Sintaxis
```js
tableRowCollectionObject.add(index, values);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Opcional. Especifica la posición relativa de la nueva fila. Si es NULL o -1, se produce la adición al final. Las filas situadas debajo de la fila insertada se desplazan hacia abajo. Indizado con cero.|
|values|(boolean or string or number)[][]|Opcional. Matriz bidimensional de valores sin formato de la fila de la tabla.|

#### <a name="returns"></a>Valores devueltos
[TableRow](tablerow.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var values = [["Sample", "Values", "For", "New", "Row"]];
    var row = tables.getItem("Table1").rows.add(null, values);
    row.load('index');
    return ctx.sync().then(function() {
        console.log(row.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getcount"></a>getCount()
Obtiene el número de filas de la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableRowCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
entero

### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtiene una fila basada en su posición en la colección.

#### <a name="syntax"></a>Sintaxis
```js
tableRowCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[TableRow](tablerow.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var tablerow = ctx.workbook.tables.getItem('Table1').rows.getItemAt(0);
    tablerow.load('name');
    return ctx.sync().then(function() {
            console.log(tablerow.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

```js
Excel.run(function (ctx) { 
    var tablerows = ctx.workbook.tables.getItem('Table1').rows;
    tablerows.load('items');
    return ctx.sync().then(function() {
        console.log("tablerows Count: " + tablerows.count);
        for (var i = 0; i < tablerows.items.length; i++)
        {
            console.log(tablerows.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```