# <a name="pageview-object-javascript-api-for-visio"></a>Objeto PageView (API de JavaScript para Visio)

Se aplica a: _Visio Online_
>**Nota:** Las API de JavaScript para Visio están actualmente en la vista previa y están sujetas a cambios. Las API de JavaScript para Visio no se admiten actualmente para su uso en entornos de producción.

Representa la clase de vista de página.

## <a name="properties"></a>Propiedades

Ninguno

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|zoom|[Doble](double.md)|Nivel de Zoom de GetSet Page.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-zoom)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|[centerViewportOnShape(ShapeId: número)](#centerviewportonshapeshapeid-number)|vacío|Aplica panorámica en el dibujo de Visio para colocar la forma especificada en el centro de la vista.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-centerViewportOnShape)|
|[fitToWindow()](#fittowindow)|vacío|Ajustar página a la ventana actual.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-fitToWindow)|
|[isShapeInViewport(Shape: Forma](#isshapeinviewportshape-shape)|bool|Para comprobar si la forma está en vista de la página o no.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-isShapeInViewport)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-load)|

## <a name="method-details"></a>Detalles del método


### <a name="centerviewportonshapeshapeid-number"></a>centerViewportOnShape(ShapeId: número)
Aplica panorámica en el dibujo de Visio para colocar la forma especificada en el centro de la vista.

#### <a name="syntax"></a>Sintaxis
```js
pageViewObject.centerViewportOnShape(ShapeId);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|ShapeId|número|ShapeId para que se vea en el centro.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    activePage.view.centerViewportOnShape(shape.Id);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="fittowindow"></a>fitToWindow()
Ajustar página a la ventana actual.

#### <a name="syntax"></a>Sintaxis
```js
pageViewObject.fitToWindow();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
vacío

### <a name="isshapeinviewportshape-shape"></a>isShapeInViewport(Shape: Forma)
Para comprobar si la forma está en vista de la página o no.

#### <a name="syntax"></a>Sintaxis
```js
pageViewObject.isShapeInViewport(Shape);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|Shape|Shape|Forma que se va a comprobar.|

#### <a name="returns"></a>Valores devueltos
bool

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
vacío

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|Posición|Posición|Objeto de posición que especifica la nueva posición de la página en la vista.|

#### <a name="returns"></a>Valores devueltos
void
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    activePage.view.zoom = 300;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

