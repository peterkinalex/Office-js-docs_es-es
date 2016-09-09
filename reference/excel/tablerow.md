# Objeto TableRow (API de JavaScript para Excel)

Representa una fila de una tabla.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|index|int|Devuelve el número de índice de la fila dentro de la colección de filas de la tabla. Indizado con cero. Solo lectura.|
|values|object[][]|Representa los valores sin formato del intervalo especificado. Los datos devueltos pueden ser de tipo cadena, número o booleano. La celda que contenga un error devolverá la cadena de error.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Elimina la fila de la tabla.|
|[getRange()](#getrange)|[Range](range.md)|Devuelve el objeto de intervalo asociado a toda la fila.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método


### delete()
Elimina la fila de la tabla.

#### Sintaxis
```js
tableRowObject.delete();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
    row.delete();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getRange()
Devuelve el objeto de intervalo asociado a toda la fila.

#### Sintaxis
```js
tableRowObject.getRange();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(0);
    var rowRange = row.getRange();
    rowRange.load('address');
    return ctx.sync().then(function() {
        console.log(rowRange.address);
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
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).tableRows.getItem(0);
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

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var newValues = [["New", "Values", "For", "New", "Row"]];
    var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
    row.values = newValues;
    row.load('values');
    return ctx.sync().then(function() {
        console.log(row.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
