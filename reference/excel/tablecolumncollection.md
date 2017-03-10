# <a name="tablecolumncollection-object-javascript-api-for-excel"></a>Objeto TableColumnCollection (API de JavaScript para Excel)

Representa una colección de todas las columnas que forman parte de la tabla.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|count|int|Devuelve el número de columnas de la tabla. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[TableColumn[]](tablecolumn.md)|Colección de objetos tableColumn. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: (boolean or string or number)[][], name: string)](#addindex-number-values-boolean-or-string-or-number-name-string)|[TableColumn](tablecolumn.md)|Agrega una nueva columna a la tabla.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|entero|Obtiene el número de columnas de la tabla.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|Obtiene un objeto de columna por nombre o identificador.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|Obtiene una columna basada en su posición en la colección.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: number or string)](#getitemornullobjectkey-number-or-string)|[TableColumn](tablecolumn.md)|Obtiene un objeto de columna por nombre o identificador. Si la columna no existe, devolverá un objeto NULL.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="addindex-number-values-boolean-or-string-or-number-name-string"></a>add(index: number, values: (boolean or string or number)[][], name: string)
Agrega una nueva columna a la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableColumnCollectionObject.add(index, values, name);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Opcional. Especifica la posición relativa de la nueva columna. Si es NULL o -1, la columna se agrega al final. Las columnas con un índice más alto se desplazarán a un lado. Indizado con cero.|
|values|(boolean or string or number)[][]|Opcional. Matriz bidimensional de valores sin formato de la columna de la tabla.|
|name|string|Opcional. Especifica el nombre de la nueva columna. Si es NULL, se utilizará el nombre predeterminado.|

#### <a name="returns"></a>Valores devueltos
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = tables.getItem("Table1").columns.add(null, values);
    column.load('name');
    return ctx.sync().then(function() {
        console.log(column.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcount"></a>getCount()
Obtiene el número de columnas de la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableColumnCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
int

### <a name="getitemkey-number-or-string"></a>getItem(key: number or string)
Obtiene un objeto de columna por nombre o identificador.

#### <a name="syntax"></a>Sintaxis
```js
tableColumnCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|Key|número o cadena| Nombre o identificador de columna.|

#### <a name="returns"></a>Valores devueltos
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem('Table1').columns.getItem(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
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
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtiene una columna basada en su posición en la colección.

#### <a name="syntax"></a>Sintaxis
```js
tableColumnCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitemornullobjectkey-number-or-string"></a>getItemOrNullObject(key: number or string)
Obtiene un objeto de columna por nombre o identificador. Si la columna no existe, devolverá un objeto NULL.

#### <a name="syntax"></a>Sintaxis
```js
tableColumnCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|Key|número o cadena| Nombre o identificador de columna.|

#### <a name="returns"></a>Valores devueltos
[TableColumn](tablecolumn.md)
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

```js
Excel.run(function (ctx) { 
    var tablecolumns = ctx.workbook.tables.getItem('Table1').columns;
    tablecolumns.load('items');
    return ctx.sync().then(function() {
        console.log("tablecolumns Count: " + tablecolumns.count);
        for (var i = 0; i < tablecolumns.items.length; i++)
        {
            console.log(tablecolumns.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```