# Objeto Workbook (API de JavaScript para Excel)

Workbook es el objeto de nivel superior que contiene los objetos de libro relacionados, como hojas de cálculo, tablas, intervalos, etc.

## Propiedades

Ninguno

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|aplicación|[Aplicación](application.md)|Representa una instancia de aplicación de Excel que contiene este libro. Solo lectura.|
|bindings|[BindingCollection](bindingcollection.md)|Representa una colección de enlaces que forman parte del libro. Solo lectura.|
|functions|[Funciones](functions.md)|Representa una instancia de aplicación de Excel que contiene este libro. Solo lectura.|
|names|[NamedItemCollection](nameditemcollection.md)|Representa una colección de elementos con nombre en el ámbito del libro (intervalos y constantes con nombre). Solo lectura.|
|tablas|[TableCollection](tablecollection.md)|Representa una colección de tablas asociadas con el libro. Solo lectura.|
|Worksheets|[WorksheetCollection](worksheetcollection.md)|Representa una colección de hojas de cálculo asociadas con el libro. Solo lectura.|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Obtiene el intervalo seleccionado actualmente en el libro.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método


### getSelectedRange()
Obtiene el intervalo seleccionado actualmente en el libro.

#### Sintaxis
```js
workbookObject.getSelectedRange();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Range](range.md)

#### Ejemplos

```js
Excel.run(function (ctx) { 
    var selectedRange = ctx.workbook.getSelectedRange();
    selectedRange.load('address');
    return ctx.sync().then(function() {
            console.log(selectedRange.address);
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
