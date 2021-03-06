# <a name="worksheetcollection-object-javascript-api-for-excel"></a>Objeto WorksheetCollection (API de JavaScript para Excel)

Representa una colección de objetos de hoja de cálculo que forman parte del libro.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|elementos|[Worksheet[]](worksheet.md)|Colección de objetos de hoja de cálculo. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[add(name: string)](#addname-string)|[Worksheet](worksheet.md)|Agrega una nueva hoja al libro. La hoja de cálculo se agregará al final de las hojas de cálculo existentes. Si desea activar la hoja de cálculo recién agregada, llame en ella a ".activate().|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getActiveWorksheet()](#getactiveworksheet)|[Worksheet](worksheet.md)|Obtiene la hoja de cálculo activa en estos momentos del libro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount(visibleOnly: bool)](#getcountvisibleonly-bool)|entero|Obtiene el número de hojas de cálculo de la colección.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Worksheet](worksheet.md)|Obtiene un objeto de hoja de cálculo mediante su nombre o identificador.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: cadena)](#getitemornullobjectkey-string)|[Worksheet](worksheet.md)|Obtiene un objeto de hoja de cálculo mediante su nombre o identificador. Si la hoja de cálculo no existe, devolverá un objeto NULL.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="addname-string"></a>add(name: string)
Agrega una nueva hoja al libro. La hoja de cálculo se agregará al final de las hojas de cálculo existentes. Si desea activar la hoja de cálculo recién agregada, llame en ella a ".activate().

#### <a name="syntax"></a>Sintaxis
```js
worksheetCollectionObject.add(name);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Opcional. Nombre de la hoja de cálculo que se va a agregar. Si se especifica, el nombre debe ser único. Si no se especifica, Excel determina el nombre de la nueva hoja de cálculo.|

#### <a name="returns"></a>Valores devueltos
[Worksheet](worksheet.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sample Name';
    var worksheet = ctx.workbook.worksheets.add(wSheetName);
    worksheet.load('name');
    return ctx.sync().then(function() {
        console.log(worksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getactiveworksheet"></a>getActiveWorksheet()
Obtiene la hoja de cálculo activa del libro.

#### <a name="syntax"></a>Sintaxis
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Worksheet](worksheet.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) {  
    var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
    activeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(activeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcountvisibleonly-bool"></a>getCount(visibleOnly: bool)
Obtiene el número de hojas de cálculo de la colección.

#### <a name="syntax"></a>Sintaxis
```js
worksheetCollectionObject.getCount(visibleOnly);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|visibleOnly|bool|Opcional. Devuelve solo las hojas de cálculo visibles si está establecido en "true". |

#### <a name="returns"></a>Valores devueltos
int

### <a name="getitemkey-string"></a>getItem(key: string)
Obtiene un objeto de hoja de cálculo mediante su nombre o identificador.

#### <a name="syntax"></a>Sintaxis
```js
worksheetCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|Key|string|Nombre o identificador de la hoja de cálculo.|

#### <a name="returns"></a>Valores devueltos
[Worksheet](worksheet.md)

### <a name="getitemornullobjectkey-string"></a>getItemOrNullObject(key: string)
Obtiene un objeto de hoja de cálculo mediante su nombre o identificador. Si la hoja de cálculo no existe, devolverá un objeto NULL.

#### <a name="syntax"></a>Sintaxis
```js
worksheetCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|Key|string|Nombre o identificador de la hoja de cálculo.|

#### <a name="returns"></a>Valores devueltos
[Worksheet](worksheet.md)
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
```js
Excel.run(function (ctx) { 
    var worksheets = ctx.workbook.worksheets;
    worksheets.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < worksheets.items.length; i++)
        {
            console.log(worksheets.items[i].name);
            console.log(worksheets.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
