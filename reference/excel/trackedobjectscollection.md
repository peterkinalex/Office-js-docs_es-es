# Objeto TrackedObjectsCollection (API de JavaScript para Office 2016)

Permite que los complementos administren las referencias de objeto de intervalo en lotes sync(). Normalmente, Excel.run() permite mantener automáticamente las referencias entre los lotes, sin tener que realizar un seguimiento de manera explícita. Sin embargo, si un escenario de complemento requiere que se realice un seguimiento de un objeto de intervalo y que se ajuste manualmente para reflejar el estado actual del intervalo de Excel subyacente, esta colección podría usarse para marcar dichos objetos para su seguimiento. Tenga en cuenta que si un objeto de intervalo está marcado para su seguimiento, debe quitarse explícitamente tras su uso para liberar memoria en Excel, incluso en caso de error.

## Propiedades
Ninguna.

## Relaciones

Ninguno

## Métodos

El objeto trackedObjectsCollection tiene definidos los métodos siguientes:

| Método     | Tipo de valor devuelto    |Descripción|
|:-----------------|:--------|:----------|
|[add(rangeObject: Range)](#addrangeobject-range)| Null             |Crea una referencia nueva en un intervalo.|
|[remove(rangeObject: Range)](#removerangeobject-range)| Null             |Elimina una referencia en el intervalo.  |
|[removeAll()](#removeall)| Null|Quita todas las referencias creadas por el complemento en el dispositivo.|


## Especificación de API 

### add(rangeObject: range)
Agrega un objeto de intervalo a trackedObjectsCollection. Se realizará un seguimiento de los cambios subyacentes en todas las solicitudes de lote y se aplicarán todas las actualizaciones de seguimiento al estado actual del objeto de intervalo. 

#### Sintaxis
```js
trackedObjectsCollection.add(rangeObject);
```

#### Parámetros

Parámetro       | Tipo   | Descripción
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| Objeto de intervalo que debe agregarse a trackedObjectCollection.

#### Valores devueltos
Null

#### Ejemplos

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    return ctx.sync(); 
});
```


### remove(rangeObject: range)

Quita un objeto de referencia de la colección. De este modo se liberan la memoria y los recursos necesarios para mantener el estado del objeto del que se realiza el seguimiento. Tenga en cuenta que si un objeto de intervalo está marcado para su seguimiento, debe quitarse explícitamente incluso en caso de error.

#### Sintaxis
```js
trackedObjectsCollection.remove(rangeObject);
```

#### Parámetros

Parámetro       | Tipo   | Descripción
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| Objeto de intervalo que debe quitarse de trackedObjectCollection.

#### Valores devueltos
Null

#### Ejemplos


```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    ctx.trackedObjectsCollection.remove(range); 
    return ctx.sync(); 
});
```

### removeAll(rangeObject: range)

Quita todas las referencias creadas por el complemento en el dispositivo.

#### Sintaxis
```js
trackedObjectsCollection.removeAll();
```

#### Parámetros

Ninguno

#### Valores devueltos
Null

#### Ejemplos

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:B2";
    var ctx = new Excel.RequestContext();
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    ctx.trackedObjectsCollection.add(range);
    ctx.load(range);
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    ctx.trackedObjectsCollection.removeAll(); 
    return ctx.sync(); 
});
```
