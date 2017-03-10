# <a name="tablecolumn-object-javascript-api-for-excel"></a>Objeto TableColumn (API de JavaScript para Excel)

Representa una columna en una tabla.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|id|int|Devuelve una clave única que identifica la columna de la tabla. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|index|int|Devuelve el número de índice de la columna dentro de la colección de columnas de la tabla. Indizado con cero. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Representa el nombre de la columna de la tabla.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|Representa los valores sin formato del rango especificado. Los datos devueltos pueden ser de tipo string, number o boolean. La celda que contenga un error devolverá la cadena de error.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|filtro|[Filter](filter.md)|Recupera el filtro aplicado a la columna. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Elimina la columna de la tabla.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado al cuerpo de datos de la columna.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado a la fila de encabezado de la columna.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado a toda la columna.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado a la fila de totales de la columna.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="delete"></a>delete()
Elimina la columna de la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableColumnObject.delete();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(2);
    column.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getdatabodyrange"></a>getDataBodyRange()
Obtiene el objeto de intervalo asociado al cuerpo de datos de la columna.

#### <a name="syntax"></a>Sintaxis
```js
tableColumnObject.getDataBodyRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var dataBodyRange = column.getDataBodyRange();
    dataBodyRange.load('address');
    return ctx.sync().then(function() {
        console.log(dataBodyRange.address);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getheaderrowrange"></a>getHeaderRowRange()
Obtiene el objeto de intervalo asociado a la fila de encabezado de la columna.

#### <a name="syntax"></a>Sintaxis
```js
tableColumnObject.getHeaderRowRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var headerRowRange = columns.getHeaderRowRange();
    headerRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(headerRowRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getrange"></a>getRange()
Obtiene el objeto de intervalo asociado a toda la columna.

#### <a name="syntax"></a>Sintaxis
```js
tableColumnObject.getRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var columnRange = columns.getRange();
    columnRange.load('address');
    return ctx.sync().then(function() {
        console.log(columnRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettotalrowrange"></a>getTotalRowRange()
Obtiene el objeto de intervalo asociado a la fila de totales de la columna.

#### <a name="syntax"></a>Sintaxis
```js
tableColumnObject.getTotalRowRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var totalRowRange = columns.getTotalRowRange();
    totalRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(totalRowRange.address);
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
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).columns.getItem(0);
    column.load('index');
    return ctx.sync().then(function() {
        console.log(column.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var tables = ctx.workbook.tables;
    var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(2);
    column.values = newValues;
    column.load('values');
    return ctx.sync().then(function() {
        console.log(column.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```