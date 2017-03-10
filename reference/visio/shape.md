# <a name="shape-object-javascript-api-for-visio"></a>Objeto de forma (API de JavaScript para Visio)

Se aplica a: _Visio Online_

Representa la clase Shape.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción|
|:---------------|:--------|:----------|
|id|int|Identificador de la forma. Solo lectura.|
|name|string|Nombre de la forma. Solo lectura.|
|select|bool|Devuelve True, si la forma está seleccionada. Usuario puede establecer True para seleccionar la forma explícitamente.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-select)|
|text|string|Texto de la forma. Solo lectura.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción|
|:---------------|:--------|:----------|
|comentarios|[CommentCollection](commentcollection.md)|Devuelve la colección de comentarios. Solo lectura.|
|hipervínculos|[HyperlinkCollection](hyperlinkcollection.md)|Devuelve la colección hipervínculos para un objeto de forma. Solo lectura.|
|shapeDataItems|[ShapeDataItemCollection](shapedataitemcollection.md)|Devuelve la sección de datos de la forma. Solo lectura.|
|subShapes|[ShapeCollection](shapecollection.md)|Obtiene la colección de subformas. Solo lectura.|
|vista|[ShapeView](shapeview.md)|Devuelve la vista de la forma. Solo lectura.|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[getBounds()](#getbounds)|[BoundingBox](boundingbox.md)|Devuelve el objeto BoundingBox que especifica el cuadro de límite de la forma.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## <a name="method-details"></a>Detalles del método


### <a name="getbounds"></a>getBounds()
Devuelve el objeto BoundingBox que especifica el cuadro de límite de la forma.

#### <a name="syntax"></a>Sintaxis
```js
shapeObject.getBounds();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[BoundingBox](boundingbox.md)

### <a name="loadparam-object"></a>load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
object.load(param);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
void
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Sample Name";
    var shape = activePage.shapes.getItem(shapeName);
    shape.load();
    return ctx.sync().then(function () {
        console.log(shape.name );
        console.log(shape.id );
        console.log(shape.Text );
        console.log(shape.Select );
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    shape.view.highlight = { color: "#E7E7E7", width: 100 };
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
