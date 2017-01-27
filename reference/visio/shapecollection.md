# <a name="shapecollection-object-javascript-api-for-visio"></a>Objeto ShapeCollection (API de JavaScript para Visio)

Se aplica a: _Visio Online_

>**Nota:** Las API de JavaScript para Visio no están disponibles actualmente para su uso en entornos de producción o de versión preliminar.

Representa la colección de formas.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|elementos|[Shape[]](shape.md)|Una colección de objetos de forma. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-items)|

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|Obtiene el número de formas de una colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-getCount)|
|[getItem(clave: número o cadena)](#getitemkey-number-or-string)|[Forma](shape.md)|Obtiene una forma mediante su clave (nombre o índice).|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-load)|

## <a name="method-details"></a>Detalles del método


### <a name="getcount"></a>getCount()
Obtiene el número de formas de una colección.

#### <a name="syntax"></a>Sintaxis
```js
shapeCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
int

#### <a name="examples"></a>Ejemplos
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var numShapesActivePage = activePage.shapes.getCount();
    return ctx.sync().then(function () {
        console.log("Shapes Count: " + numShapesActivePage.value);
    });

}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitemkey-number-or-string"></a>getItem(clave: número o cadena)
Obtiene una forma mediante su clave (nombre o índice).

#### <a name="syntax"></a>Sintaxis
```js
shapeCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|Key|número o cadena|Clave es el nombre o el índice de la forma que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[Forma](shape.md)

### <a name="loadparam-object"></a>load(param: objeto)
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
