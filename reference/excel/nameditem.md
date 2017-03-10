# <a name="nameditem-object-javascript-api-for-excel"></a>Objeto NamedItem (API de JavaScript para Excel)

Representa un nombre definido para un rango de celdas o un valor. Los nombres pueden ser objetos primitivos con nombre (como puede verse en el tipo siguiente), un objeto de rango o una referencia a un rango. Este objeto puede usarse para obtener un objeto de rango asociado a nombres.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|comment|string|Representa el comentario asociado a este nombre.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Nombre del objeto. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|scope|string|Indica si el nombre está en el ámbito del libro o de una hoja de cálculo específica. Solo lectura. Los valores posibles son: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Indica el tipo del valor que devuelve la fórmula del nombre. Solo lectura. Los valores posibles son: String, Integer, Double, Boolean, Range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|value|objeto|Representa el valor calculado por la fórmula del nombre. Para un rango con nombre, devolverá la dirección del rango. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Especifica si el objeto es visible o no.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|worksheet|[Worksheet](worksheet.md)|Devuelve la hoja de cálculo que tiene como ámbito el elemento con nombre. Se produce un error si el ámbito del elemento es el libro. Solo lectura.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetOrNullObject|[Worksheet](worksheet.md)|Devuelve la hoja de cálculo que tiene como ámbito el elemento con nombre. Devuelve un objeto NULL si el ámbito del elemento es el libro. Solo lectura.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Elimina el nombre especificado.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Devuelve el objeto de rango asociado al nombre. Se produce un error si el tipo del elemento con nombre no es un rango.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRangeOrNullObject()](#getrangeornullobject)|[Range](range.md)|Devuelve el objeto de rango asociado al nombre. Devuelve un objeto NULL si el tipo de elemento con nombre no es un rango.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="delete"></a>delete()
Elimina el nombre especificado.

#### <a name="syntax"></a>Sintaxis
```js
namedItemObject.delete();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

### <a name="getrange"></a>getRange()
Devuelve el objeto de rango asociado al nombre. Se produce un error si el tipo del elemento con nombre no es un rango.

#### <a name="syntax"></a>Sintaxis
```js
namedItemObject.getRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos

Devuelve el objeto Range que está asociado al nombre. `null` si el nombre no es del tipo `Range`. Nota: Esta API actualmente solo admite elementos en el ámbito del libro.**

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


### <a name="getrangeornullobject"></a>getRangeOrNullObject()
Devuelve el objeto de rango asociado al nombre. Devuelve un objeto NULL si el tipo de elemento con nombre no es un rango.

#### <a name="syntax"></a>Sintaxis
```js
namedItemObject.getRangeOrNullObject();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

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
