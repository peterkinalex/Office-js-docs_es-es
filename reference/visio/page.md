# <a name="page-object-javascript-api-for-visio"></a>Objeto Page (API de JavaScript para Visio)

Se aplica a: _Visio Online_
>**Nota:** Las API de JavaScript para Visio están actualmente en la vista previa y están sujetas a cambios. Las API de JavaScript para Visio no se admiten actualmente para su uso en entornos de producción.

Representa la clase de página.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|índice|int|Índice de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-index)|
|isBackground|bool|Si la página es una página de fondo o no. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-isBackground)|
|name|string|Nombre de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-name)|

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|formas|[ShapeCollection](shapecollection.md)|Formas en la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-shapes)|
|vista|[PageView](pageview.md)|Devuelve la vista de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-view)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|[activate()](#activate)|vacío|Establece la página como la página activa del documento.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-activate)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-load)|

## <a name="method-details"></a>Detalles del método


### <a name="activate"></a>activate()
Establece la página como la página activa del documento.

#### <a name="syntax"></a>Sintaxis
```js
pageObject.activate();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

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
