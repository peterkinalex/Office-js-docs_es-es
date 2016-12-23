# <a name="shape-object-javascript-api-for-visio"></a>Objeto de forma (API de JavaScript para Visio)

Se aplica a: _Visio Online_
>**Nota:** Las API de JavaScript para Visio están actualmente en la vista previa y están sujetas a cambios. Las API de JavaScript para Visio no se admiten actualmente para su uso en entornos de producción.

Representa la clase de forma.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|id|int|Identificador de la forma. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-id)|
|name|string|Nombre de la forma. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-name)|
|seleccionar|bool|Devuelve True, si la forma está seleccionada. Usuario puede establecer True para seleccionar la forma explícitamente.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-select)|
|text|string|Texto de la forma. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-text)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|hipervínculos|[HyperlinkCollection](hyperlinkcollection.md)|Devuelve la colección hipervínculos para un objeto de forma. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-hyperlinks)|
|shapeDataItems|[ShapeDataItemCollection](shapedataitemcollection.md)|Devuelve la sección de datos de la forma. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-shapeDataItems)|
|subShapes|[ShapeCollection](shapecollection.md)|Obtiene la colección de subformas. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-subShapes)|
|vista|[ShapeView](shapeview.md)|Devuelve la vista de la forma. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-view)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-load)|

## <a name="method-details"></a>Detalles del método

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