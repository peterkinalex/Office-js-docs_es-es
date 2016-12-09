# <a name="binding-object-javascript-api-for-excel"></a>Objeto Binding (API de JavaScript para Excel)

Representa un enlace de Office.js que se define en el libro.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|id|string|Representa el identificador de enlace. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Devuelve el tipo de enlace. Solo lectura. Los valores posibles son: Range, Table, Text.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Elimina el enlace.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Devuelve el intervalo representado por el enlace. Se producirá un error si el enlace no es del tipo correcto.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTable()](#gettable)|[Table](table.md)|Devuelve la tabla representada por el enlace. Se producirá un error si el enlace no es del tipo correcto.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getText()](#gettext)|string|Devuelve el texto representado por el enlace. Se producirá un error si el enlace no es del tipo correcto.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy que se ha creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="delete"></a>delete()
Elimina el enlace.

#### <a name="syntax"></a>Sintaxis
```js
bindingObject.delete();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

### <a name="getrange"></a>getRange()
Devuelve el intervalo representado por el enlace. Se producirá un error si el enlace no es del tipo correcto.

#### <a name="syntax"></a>Sintaxis
```js
bindingObject.getRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
El ejemplo siguiente usa el objeto de enlace para obtener el intervalo asociado.

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var range = binding.getRange();
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettable"></a>getTable()
Devuelve la tabla representada por el enlace. Se producirá un error si el enlace no es del tipo correcto.

#### <a name="syntax"></a>Sintaxis
```js
bindingObject.getTable();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Table](table.md)

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var table = binding.getTable();
    table.load('name');
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


### <a name="gettext"></a>getText()
Devuelve el texto representado por el enlace. Se producirá un error si el enlace no es del tipo correcto.

#### <a name="syntax"></a>Sintaxis
```js
bindingObject.getText();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
string

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var text = binding.getText();
    binding.load('text');
    return ctx.sync().then(function() {
        console.log(text);
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
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    binding.load('type');
    return ctx.sync().then(function() {
        console.log(binding.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
