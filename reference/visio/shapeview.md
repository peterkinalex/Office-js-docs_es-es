# <a name="shapeview-object-javascript-api-for-visio"></a>Objeto ShapeView (API de JavaScript para Visio)

Se aplica a: _Visio Online_
>**Nota:** Las API de JavaScript para Visio están actualmente en la vista previa y están sujetas a cambios. Las API de JavaScript para Visio no se admiten actualmente para su uso en entornos de producción.

Representa la clase ShapeView.

## <a name="properties"></a>Propiedades

Ninguno

## <a name="relationships"></a>Relaciones
Ninguno

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|[addOverlay(OverlayType: OverlayType, contenido: cadena, HorizontalAlignment: HorizontalAlignment, VerticalAlignment: VerticalAlignment, ancho: número, alto: número)](#addoverlayoverlaytype-overlaytype-content-string-horizontalalignment-horizontalalignment-verticalalignment-verticalalignment-width-number-height-number)|int|Agrega una superposición encima de la forma.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-addOverlay)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-load)|
|[removeOverlay(OverlayId: número)](#removeoverlayoverlayid-number)|vacío|Quita una superposición específica o todas las superposiciones de la forma.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-removeOverlay)|

## <a name="method-details"></a>Detalles del método


### <a name="addoverlayoverlaytype-overlaytype-content-string-horizontalalignment-horizontalalignment-verticalalignment-verticalalignment-width-number-height-number"></a>addOverlay(OverlayType: OverlayType, contenido: cadena, HorizontalAlignment: HorizontalAlignment, VerticalAlignment: VerticalAlignment, ancho: número, alto: número)
Agrega una superposición encima de la forma.

#### <a name="syntax"></a>Sintaxis
```js
shapeViewObject.addOverlay(OverlayType, Content, HorizontalAlignment, VerticalAlignment, Width, Height);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|OverlayType|OverlayType|Un tipo superposición: texto, imagen.|
|Contenido|string|Contenido de superposición.|
|HorizontalAlignment|HorizontalAlignment|Alineación horizontal de la superposición: izquierda, centro, derecha|
|VerticalAlignment|VerticalAlignment|Alineación vertical de superposición: superior, intermedio, inferior|
|Ancho|número|Ancho de superposición.|
|Alto|número|Alto de superposición.|

#### <a name="returns"></a>Valores devueltos
int

### <a name="loadparam-object"></a>load(param: objeto)
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
vacío

### <a name="removeoverlayoverlayid-number"></a>removeOverlay(OverlayId: número)
Quita una superposición específica o todas las superposiciones de la forma.

#### <a name="syntax"></a>Sintaxis
```js
shapeViewObject.removeOverlay(OverlayId);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|OverlayId|number|Un identificador de superposición. Elimina el identificador específico de superposición de la forma.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    var overlayId=shape.view.addOverlay(1, "Visio Online", 2, 2, 50, 50);
    return ctx.sync();
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
    shape.view.removeOverlay(1);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
