# <a name="bindingcollection-object-javascript-api-for-excel"></a>Objeto BindingCollection (API de JavaScript para Excel)

Representa la colección de todos los objetos de enlace que forman parte del libro.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|count|entero|Devuelve el número de enlaces incluidos en la colección. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Binding[]](binding.md)|Colección de objetos de enlace. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[add(range: Range or string, bindingType: string, id: string)](#addrange-range-or-string-bindingtype-string-id-string)|[Binding](binding.md)|Agregar un enlace nuevo a un intervalo determinado.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromNamedItem(name: string, bindingType: string, id: string)](#addfromnameditemname-string-bindingtype-string-id-string)|[Binding](binding.md)|Agregar un enlace nuevo basándose en un elemento con nombre del libro.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromSelection(bindingType: string, id: string)](#addfromselectionbindingtype-string-id-string)|[Binding](binding.md)|Agregar un enlace nuevo basándose en la selección actual.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|entero|Obtiene el número de enlaces de la colección.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|Obtiene un objeto de enlace por identificador.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|Obtiene un objeto de enlace según su posición en la matriz de elementos.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(id: string)](#getitemornullobjectid-string)|[Binding](binding.md)|Obtiene un objeto de enlace por identificador. Si no existe el objeto de enlace, devolverá un objeto nulo.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="addrange-range-or-string-bindingtype-string-id-string"></a>add(range: Range or string, bindingType: string, id: string)
Agregar un enlace nuevo a un intervalo determinado.

#### <a name="syntax"></a>Sintaxis
```js
bindingCollectionObject.add(range, bindingType, id);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|range|Range o string|Intervalo al que se va a vincular el enlace. Puede ser un objeto de intervalo de Excel o una cadena. Si es una cadena, debe incluir la dirección completa, incluido el nombre de la hoja|
|bindingType|string|Tipo de enlace.  Los valores posibles son: Range, Table, Text|
|id|string|Nombre del enlace.|

#### <a name="returns"></a>Valores devueltos
[Binding](binding.md)

### <a name="addfromnameditemname-string-bindingtype-string-id-string"></a>addFromNamedItem(name: string, bindingType: string, id: string)
Agregar un enlace nuevo basándose en un elemento con nombre del libro.

#### <a name="syntax"></a>Sintaxis
```js
bindingCollectionObject.addFromNamedItem(name, bindingType, id);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Nombre desde el que se va a crear el enlace.|
|bindingType|string|Tipo de enlace.  Los valores posibles son: Range, Table, Text|
|id|string|Nombre del enlace.|

#### <a name="returns"></a>Valores devueltos
[Binding](binding.md)

### <a name="addfromselectionbindingtype-string-id-string"></a>addFromSelection(bindingType: string, id: string)
Agregar un enlace nuevo basándose en la selección actual.

#### <a name="syntax"></a>Sintaxis
```js
bindingCollectionObject.addFromSelection(bindingType, id);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|bindingType|string|Tipo de enlace.  Los valores posibles son: Range, Table, Text|
|id|string|Nombre del enlace.|

#### <a name="returns"></a>Valores devueltos
[Binding](binding.md)

### <a name="getcount"></a>getCount()
Obtiene el número de enlaces de la colección.

#### <a name="syntax"></a>Sintaxis
```js
bindingCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
entero

### <a name="getitemid-string"></a>getItem(id: string)
Obtiene un objeto de enlace por identificador.

#### <a name="syntax"></a>Sintaxis
```js
bindingCollectionObject.getItem(id);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|id|string|Identificador del objeto de contenido que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[Binding](binding.md)

#### <a name="examples"></a>Ejemplos

Crear un enlace de tabla para supervisar los cambios en los datos de la tabla. Cuando se modifica algún dato, el color de fondo de la tabla cambiará a naranja.

```js
function addEventHandler() {
    //Create Table1
Excel.run(function (ctx) { 
    ctx.workbook.tables.add("Sheet1!A1:C4", true);
    return ctx.sync().then(function() {
             console.log("My Diet Data Inserted!");
    })
    .catch(function (error) {
             console.log(JSON.stringify(error));
    });
});
    //Create a new table binding for Table1
Office.context.document.bindings.addFromNamedItemAsync("Table1", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
        console.log("Action failed with error: " + asyncResult.error.message);
    }
    else {
        // If succeeded, then add event handler to the table binding.
        Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
    }
});
}
    
// when data in the table is changed, this event will be triggered.
function onBindingDataChanged(eventArgs) {
Excel.run(function (ctx) { 
    // highlight the table in orange to indicate data has been changed.
    ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
    return ctx.sync().then(function() {
            console.log("The value in this table got changed!");
    })
    .catch(function (error) {
            console.log(JSON.stringify(error));
    });
});
}

```



#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
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


### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtiene un objeto de enlace según su posición en la matriz de elementos.

#### <a name="syntax"></a>Sintaxis
```js
bindingCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[Binding](binding.md)

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
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


### <a name="getitemornullobjectid-string"></a>getItemOrNullObject(id: string)
Obtiene un objeto de enlace por identificador. Si no existe el objeto de enlace, devolverá un objeto nulo.

#### <a name="syntax"></a>Sintaxis
```js
bindingCollectionObject.getItemOrNullObject(id);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|id|string|Identificador del objeto de contenido que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[Binding](binding.md)
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < bindings.items.length; i++)
        {
            console.log(bindings.items[i].id);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Obtener el número de enlaces.

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('count');
    return ctx.sync().then(function() {
        console.log("Bindings: Count= " + bindings.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
