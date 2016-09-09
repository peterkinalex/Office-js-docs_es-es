# Objeto TableSort (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS y Office 2016_

Administra operaciones de ordenación en objetos Table.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|matchCase|bool|Indica si última ordenación de la tabla distinguía mayúsculas de minúsculas. Solo lectura.|
|method|string|Representa el método de ordenación de caracteres chinos usado por última vez para ordenar la tabla. Solo lectura. Los valores posibles son: PinYin, StrokeCount.|

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|fields|[SortField](sortfield.md)|Representa las condiciones actuales que se usaron por última vez para ordenar la tabla. Solo lectura.|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[apply(fields: SortField[], matchCase: bool, method: string)](#applyfields-sortfield-matchcase-bool-method-string)|void|Realiza una operación de ordenación.|
|[clear()](#clear)|void|Borra la ordenación que se aplica actualmente en la tabla. Aunque esto no modifica la ordenación de la tabla, borra el estado de los botones de encabezado.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|
|[reapply()](#reapply)|void|Vuelve a aplicar los parámetros de ordenación actuales a la tabla.|

## Detalles del método


### apply(fields: SortField[], matchCase: bool, method: string)
Realiza una operación de ordenación.

#### Sintaxis
```js
tableSortObject.apply(fields, matchCase, method);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|fields|SortField[]|La lista de condiciones por las que realizar la ordenación.|
|matchCase|bool|Opcional. Indica si la ordenación de cadenas distingue mayúsculas de minúsculas.|
|method|string|Opcional. Método de ordenación que se usa para los caracteres chinos.  Los valores posibles son: PinYin, StrokeCount|

#### Valores devueltos
void

#### Ejemplos
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
Borra la ordenación que se aplica actualmente en la tabla. Aunque esto no modifica la ordenación de la tabla, borra el estado de los botones de encabezado.

#### Sintaxis
```js
tableSortObject.clear();
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
    table.sort.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

### load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### Syntax
```js
object.load(param);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### Valores devueltos
void

### reapply()
Vuelve a aplicar los parámetros de ordenación actuales a la tabla.

#### Sintaxis
```js
tableSortObject.reapply();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

####Ejemplos
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.reapply();   
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});