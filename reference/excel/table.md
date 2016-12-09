# <a name="table-object-javascript-api-for-excel"></a>Objeto Table (API de JavaScript para Excel)

Representa una tabla de Excel.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|highlightFirstColumn|bool|Indica si la primera columna contiene un formato especial.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|highlightLastColumn|bool|Indica si la última columna contiene un formato especial.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|id|int|Devuelve un valor que identifica de forma única la tabla de un libro determinado. El valor del identificador permanece igual, incluso cuando se cambia el nombre de la tabla. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Nombre de la tabla.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedColumns|bool|Indica si las columnas muestran un formato con bandas en el que las columnas impares están resaltadas de manera diferente que las pares para facilitar la lectura de la tabla.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedRows|bool|Indica si las filas muestran un formato con bandas en el que las filas impares están resaltadas de manera diferente que las pares para facilitar la lectura de la tabla.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showFilterButton|bool|Indica si los botones de filtro son visibles en la parte superior de cada encabezado de columna. Esta configuración solo se permite si la tabla contiene una fila de encabezado.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showHeaders|bool|Indica si la fila de encabezado está visible o no. Este valor puede establecerse para que muestre o quite la fila de encabezado.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTotals|bool|Indica si la fila de totales está visible o no. Este valor puede establecerse para que muestre o quite la fila de totales.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|Valor constante que representa el estilo de tabla. Los valores posibles son: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. También puede especificarse un estilo personalizado definido por el usuario presente en el libro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|columns|[TableColumnCollection](tablecolumncollection.md)|Representa una colección de todas las columnas de la tabla. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rows|[TableRowCollection](tablerowcollection.md)|Representa una colección de todas las filas de la tabla. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sort|[TableSort](tablesort.md)|Representa la ordenación de la tabla. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|La hoja de cálculo que contiene la tabla actual. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[clearFilters()](#clearfilters)|void|Borra todos los filtros aplicados actualmente en la tabla.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[convertToRange()](#converttorange)|[Range](range.md)|Convierte la tabla en un rango de celdas normal. Se conservan todos los datos.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|Elimina la tabla.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado al cuerpo de datos de la tabla.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado a la fila de encabezado de la tabla.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado a toda la tabla.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado a la fila de totales de la tabla.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy que se ha creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[reapplyFilters()](#reapplyfilters)|void|Vuelve a aplicar todos los filtros aplicados actualmente en la tabla.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="clearfilters"></a>clearFilters()
Borra todos los filtros aplicados actualmente en la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableObject.clearFilters();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

### <a name="converttorange"></a>convertToRange()
Convierte la tabla en un rango de celdas normal. Se conservan todos los datos.

#### <a name="syntax"></a>Sintaxis
```js
tableObject.convertToRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.convertToRange();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="delete"></a>delete()
Elimina la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableObject.delete();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getdatabodyrange"></a>getDataBodyRange()
Obtiene el objeto de intervalo asociado al cuerpo de datos de la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableObject.getDataBodyRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableDataRange = table.getDataBodyRange();
    tableDataRange.load('address')
    return ctx.sync().then(function() {
            console.log(tableDataRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getheaderrowrange"></a>getHeaderRowRange()
Obtiene el objeto de intervalo asociado a la fila de encabezado de la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableObject.getHeaderRowRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableHeaderRange = table.getHeaderRowRange();
    tableHeaderRange.load('address');
    return ctx.sync().then(function() {
        console.log(tableHeaderRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrange"></a>getRange()
Obtiene el objeto de intervalo asociado a toda la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableObject.getRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableRange = table.getRange();
    tableRange.load('address'); 
    return ctx.sync().then(function() {
            console.log(tableRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettotalrowrange"></a>getTotalRowRange()
Obtiene el objeto de intervalo asociado a la fila de totales de la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableObject.getTotalRowRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableTotalsRange = table.getTotalRowRange();
    tableTotalsRange.load('address');   
    return ctx.sync().then(function() {
            console.log(tableTotalsRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="loadparam-object"></a>load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
object.load(param);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
void

### <a name="reapplyfilters"></a>reapplyFilters()
Vuelve a aplicar todos los filtros aplicados actualmente en la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableObject.reapplyFilters();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

Obtener una tabla por nombre. 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.load('index')
    return ctx.sync().then(function() {
            console.log(table.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Obtener una tabla por índice.

```js
Excel.run(function (ctx) { 
    var index = 0;
    var table = ctx.workbook.tables.getItemAt(0);
    table.load('id')
    return ctx.sync().then(function() {
            console.log(table.id);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Establecer el estilo de tabla. 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.name = 'Table1-Renamed';
    table.showTotals = false;
    table.style = 'TableStyleMedium2';
    table.load('tableStyle');
    return ctx.sync().then(function() {
            console.log(table.style);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
