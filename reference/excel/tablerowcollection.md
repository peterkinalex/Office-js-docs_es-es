# <a name="tablerowcollection-object-(javascript-api-for-excel)"></a>Objeto TableRowCollection (API de JavaScript para Excel)

Representa una colección de todas las filas que forman parte de la tabla.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|count|int|Devuelve el número de filas de la tabla. Solo lectura.|
|items|[TableRow[]](tablerow.md)|Colección de objetos TableRow. Solo lectura.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableRow](tablerow.md)|Agrega una nueva fila a la tabla.|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|Obtiene una fila basada en su posición en la colección.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## <a name="method-details"></a>Detalles del método


### <a name="add(index:-number,-values:-(boolean-or-string-or-number)[][])"></a>add(index: number, values: (boolean or string or number)[][])
Agrega una nueva fila a la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableRowCollectionObject.add(index, values);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Opcional. Especifica la posición relativa de la nueva fila. Si es null, se produce la adición al final. Las filas situadas debajo de la fila insertada se desplazan hacia abajo. Indexado con cero.|
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

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtiene una fila basada en su posición en la colección.

#### <a name="syntax"></a>Sintaxis
```js
tableRowCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
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