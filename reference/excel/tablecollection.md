# <a name="tablecollection-object-javascript-api-for-excel"></a>Objeto TableCollection (API de JavaScript para Excel)

Representa una colección de todas las tablas que forman parte del libro o la hoja de cálculo, dependiendo de cómo se haya alcanzado.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|count|int|Devuelve el número de tablas del libro. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Table[]](table.md)|Colección de objetos de tabla. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|[add(address: [object, hasHeaders: bool)](#addaddress-object-hasheaders-bool)|[Table](table.md)|Crea una tabla nueva. El objeto de rango o la dirección de origen determinan la hoja de cálculo a la que se agregará la tabla. Si no se puede agregar la tabla (por ejemplo, porque la dirección no es válida o porque la tabla se superpondría con otra tabla), se producirá un error.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|entero|Obtiene el número de tablas de la colección.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Table](table.md)|Obtiene una tabla por nombre o identificador.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|Obtiene una tabla basada en su posición en la colección.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: number or string)](#getitemornullobjectkey-number-or-string)|[Table](table.md)|Obtiene una tabla por nombre o identificador. Si la tabla no existe, devolverá un objeto NULL.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="addaddress-object-hasheaders-bool"></a>add(address: [object, hasHeaders: bool)
Crea una tabla nueva. El objeto de rango o la dirección de origen determinan la hoja de cálculo a la que se agregará la tabla. Si no se puede agregar la tabla (por ejemplo, porque la dirección no es válida o porque la tabla se superpondría con otra tabla), se producirá un error.

#### <a name="syntax"></a>Sintaxis
```js
tableCollectionObject.add(address, hasHeaders);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|address|[object|Objeto de rango, dirección de cadena o nombre del rango que representa el origen de datos. Si la dirección no contiene un nombre de hoja, se usa la hoja activa en ese momento. En 1.1 se utiliza el parámetro de cadena; en 1.3 se puede usar también el objeto Range.|
|hasHeaders|bool|Valor booleano que indica si los datos que se están importando tienen etiquetas de columna. Si el origen no contiene encabezados (es decir, cuando esta propiedad se establece en false), Excel generará automáticamente el encabezado desplazando los datos hacia abajo una fila.|

#### <a name="returns"></a>Valores devueltos
[Table](table.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
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

### <a name="getcount"></a>getCount()
Obtiene el número de tablas de la colección.

#### <a name="syntax"></a>Sintaxis
```js
tableCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
int

### <a name="getitemkey-number-or-string"></a>getItem(key: number or string)
Obtener una tabla por nombre o identificador.

#### <a name="syntax"></a>Sintaxis
```js
tableCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|Key|number o string|Nombre o identificador de la tabla que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[Table](table.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
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


#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
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


### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtiene una tabla basada en su posición en la colección.

#### <a name="syntax"></a>Sintaxis
```js
tableCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[Table](table.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
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


### <a name="getitemornullobjectkey-number-or-string"></a>getItemOrNullObject(key: number or string)
Obtiene una tabla por nombre o identificador. Si la tabla no existe, devolverá un objeto NULL.

#### <a name="syntax"></a>Sintaxis
```js
tableCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|Key|number o string|Nombre o identificador de la tabla que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[Table](table.md)
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load();
    return ctx.sync().then(function() {
        console.log("tables Count: " + tables.count);
        for (var i = 0; i < tables.items.length; i++)
        {
            console.log(tables.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Obtener el número de tablas.

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load('count');
    return ctx.sync().then(function() {
        console.log(tables.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```