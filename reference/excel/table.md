# Objeto Table (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS y Office 2016_

Representa una tabla de Excel.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|id|int|Devuelve un valor que identifica de forma única la tabla de un libro determinado. El valor del identificador permanece igual, incluso cuando se cambia el nombre de la tabla. Solo lectura.|
|name|string|Nombre de la tabla.|
|showHeaders|bool|Indica si la fila de encabezado está visible o no. Este valor puede establecerse para que muestre o quite la fila de encabezado.|
|showTotals|bool|Indica si la fila de totales está visible o no. Este valor puede establecerse para que muestre o quite la fila de totales.|
|style|string|Valor constante que representa el estilo de tabla. Los valores posibles son: de TableStyleLight1 a TableStyleLight21, de TableStyleMedium1 a TableStyleMedium28, de TableStyleStyleDark1 a TableStyleStyleDark11. También puede especificarse un estilo personalizado definido por el usuario presente en el libro.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|columnas|[TableColumnCollection](tablecolumncollection.md)|Representa una colección de todas las columnas de la tabla. Solo lectura.|
|Rows|[TableRowCollection](tablerowcollection.md)|Representa una colección de todas las filas de la tabla. Solo lectura.|
|sort|[TableSort](tablesort.md)|Representa la configuración de ordenación de la tabla. Solo lectura.|
|worksheet|[Hoja de cálculo](worksheet.md)|Hoja de cálculo que contiene la tabla actual. Solo lectura.|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[clearFilters()](#clearfilters)|void|Borra todos los filtros aplicados actualmente en la tabla.|
|[convertToRange()](#converttorange)|[Range](range.md)|Convierte la tabla en un rango de celdas normal. Se conservan todos los datos.|
|[delete()](#delete)|void|Elimina la tabla.|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado al cuerpo de datos de la tabla.|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado a la fila de encabezado de la tabla.|
|[getRange()](#getrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado a toda la tabla.|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Obtiene el objeto de intervalo asociado a la fila de totales de la tabla.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|
|[reapplyFilters()](#reapplyfilters)|void|Vuelve a aplicar todos los filtros aplicados actualmente en la tabla.|

## Detalles del método


### clearFilters()
Borra todos los filtros aplicados actualmente en la tabla.

#### Sintaxis
```js
tableObject.clearFilters();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

### convertToRange()
Convierte la tabla en un rango de celdas normal. Se conservan todos los datos.

#### Sintaxis
```js
tableObject.convertToRange();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos
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

### delete()
Elimina la tabla.

#### Sintaxis
```js
tableObject.delete();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos
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


### getDataBodyRange()
Obtiene el objeto de intervalo asociado al cuerpo de datos de la tabla.

#### Sintaxis
```js
tableObject.getDataBodyRange();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos
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

### getHeaderRowRange()
Obtiene el objeto de intervalo asociado a la fila de encabezado de la tabla.

#### Sintaxis
```js
tableObject.getHeaderRowRange();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos
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


### getRange()
Obtiene el objeto de intervalo asociado a toda la tabla.

#### Sintaxis
```js
tableObject.getRange();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos
```js
Excel.run(function (ctx) { 
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


### getTotalRowRange()
Obtiene el objeto de intervalo asociado a la fila de totales de la tabla.

#### Sintaxis
```js
tableObject.getTotalRowRange();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos
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
    table.name('name')
    return ctx.sync().then(function() {
            console.log(table.name);
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
    table.tableStyle = 'TableStyleMedium2';
    table.load('tableStyle');
    return ctx.sync().then(function() {
            console.log(table.tableStyle);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
