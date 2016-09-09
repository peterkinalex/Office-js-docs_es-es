# Objeto NamedItem (API de JavaScript para Excel)

Representa un nombre definido para un intervalo de celdas o un valor. Los nombres pueden ser objetos primitivos con nombre (como puede verse en el tipo siguiente), un objeto de intervalo o una referencia a un intervalo. Este objeto puede usarse para obtener un objeto de intervalo asociado a nombres.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|name|string|Nombre del objeto. Solo lectura.|
|type|string|Indica el tipo de referencia que está asociado al nombre. Solo lectura. Los valores posibles son: String, Integer, Double, Boolean, Range.|
|value|object|Representa la fórmula a la que tiene que hacer referencia el nombre, según su definición (por ejemplo, =Hoja14!$B$2:$H$12, =4,75, etc.). Solo lectura.|
|visible|bool|Especifica si el objeto está visible o no.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|Devuelve el objeto de intervalo asociado al nombre. Produce una excepción si el tipo del elemento con nombre no es un intervalo.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método


### getRange()
Devuelve el objeto de intervalo asociado al nombre. Produce una excepción si el tipo del elemento con nombre no es un intervalo.

#### Sintaxis
```js
namedItemObject.getRange();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos

Devuelve el objeto de intervalo que está asociado al nombre. `null` si el nombre no es del tipo `Range`. Nota: Esta API actualmente solo admite elementos del ámbito del libro.

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var range = names.getItem('MyRange').getRange();
    range.load('address');
    return ctx.sync().then(function() {
            console.log(range.address);
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
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    namedItem.load('type');
    return ctx.sync().then(function() {
            console.log(namedItem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
