# <a name="workbook-object-(javascript-api-for-excel)"></a>Objeto Workbook (API de JavaScript para Excel)

Workbook es el objeto de nivel superior que contiene los objetos de libro relacionados, como hojas de cálculo, tablas, intervalos, etc.

## <a name="properties"></a>Propiedades

Ninguno

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|application|[Application](application.md)|Representa una instancia de aplicación de Excel que contiene este libro. Solo lectura.|
|bindings|[BindingCollection](bindingcollection.md)|Representa una colección de enlaces que forman parte del libro. Solo lectura.|
|functions|[Functions](functions.md)|Representa una instancia de aplicación de Excel que contiene este libro. Solo lectura.|
|names|[NamedItemCollection](nameditemcollection.md)|Representa una colección de elementos con nombre en el ámbito del libro (intervalos y constantes con nombre). Solo lectura.|
|tables|[TableCollection](tablecollection.md)|Representa una colección de tablas asociadas con el libro. Solo lectura.|
|worksheets|[WorksheetCollection](worksheetcollection.md)|Representa una colección de hojas de cálculo asociadas con el libro. Solo lectura.|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Obtiene el intervalo seleccionado actualmente en el libro.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## <a name="method-details"></a>Detalles del método


### <a name="getselectedrange()"></a>getSelectedRange()
Obtiene el intervalo seleccionado actualmente en el libro.

#### <a name="syntax"></a>Sintaxis
```js
workbookObject.getSelectedRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

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
