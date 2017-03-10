# <a name="nameditemcollection-object-javascript-api-for-excel"></a>Objeto NamedItemCollection (API de JavaScript para Excel)

Una colección de todos los objetos namedItem que forman parte del libro o la hoja de cálculo, dependiendo de cómo se haya alcanzado.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|elementos|[NamedItem[]](nameditem.md)|Colección de objetos namedItem. Solo lectura.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|[add(name: string, reference: Range or string, comment: string)](#addname-string-reference-range-or-string-comment-string)|[NamedItem](nameditem.md)|Agrega un nuevo nombre a la colección del ámbito especificado.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[addFormulaLocal(name: string, formula: string, comment: string)](#addformulalocalname-string-formula-string-comment-string)|[NamedItem](nameditem.md)|Agrega un nuevo nombre a la colección del ámbito especificado, empleando la configuración regional del usuario para la fórmula.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|entero|Obtiene el número de elementos con nombre de la colección.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Obtiene un objeto NamedItem mediante su nombre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[NamedItem](nameditem.md)|Obtiene un objeto NamedItem mediante su nombre. Si no existe el objeto NamedItem, devolverá un objeto NULL.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="addname-string-reference-range-or-string-comment-string"></a>add(name: string, reference: Range or string, comment: string)
Agrega un nuevo nombre a la colección del ámbito especificado.

#### <a name="syntax"></a>Sintaxis
```js
namedItemCollectionObject.add(name, reference, comment);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Nombre del elemento con nombre.|
|reference|intervalo o cadena|Fórmula o rango a los que se refiere el nombre.|
|comment|string|Opcional. Comentario asociado al elemento con nombre|

#### <a name="returns"></a>Valores devueltos
[NamedItem](nameditem.md)

### <a name="addformulalocalname-string-formula-string-comment-string"></a>addFormulaLocal(name: string, formula: string, comment: string)
Agrega un nuevo nombre a la colección del ámbito especificado, empleando la configuración regional del usuario para la fórmula.

#### <a name="syntax"></a>Sintaxis
```js
namedItemCollectionObject.addFormulaLocal(name, formula, comment);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Elemento "name" del elemento con nombre.|
|formula|string|Fórmula de la configuración regional del usuario a la que se refiere el nombre.|
|comment|string|Opcional. Comentario asociado al elemento con nombre|

#### <a name="returns"></a>Valores devueltos
[NamedItem](nameditem.md)

### <a name="getcount"></a>getCount()
Obtiene el número de elementos con nombre de la colección.

#### <a name="syntax"></a>Sintaxis
```js
namedItemCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
entero

### <a name="getitemname-string"></a>getItem(name: string)
Obtiene un objeto namedItem mediante su nombre.

#### <a name="syntax"></a>Sintaxis
```js
namedItemCollectionObject.getItem(name);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Nombre de namedItem.|

#### <a name="returns"></a>Valores devueltos
[NamedItem](nameditem.md)

#### <a name="examples"></a>Ejemplos

```js
Excel.run(function (ctx) { 
    var sheetName = 'Sheet1';
    var nameditem = ctx.workbook.names.getItem(sheetName);
    nameditem.load('type');
    return ctx.sync().then(function() {
            console.log(nameditem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
Obtiene un objeto NamedItem mediante su nombre. Si no existe el objeto NamedItem, devolverá un objeto NULL.

#### <a name="syntax"></a>Sintaxis
```js
namedItemCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Nombre de namedItem.|

#### <a name="returns"></a>Valores devueltos
[NamedItem](nameditem.md)
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

```js
Excel.run(function (ctx) { 
    var nameditems = ctx.workbook.names;
    nameditems.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < nameditems.items.length; i++)
        {
            console.log(nameditems.items[i].name);
            console.log(nameditems.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


