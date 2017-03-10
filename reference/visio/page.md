# <a name="page-object-javascript-api-for-visio"></a>Objeto Page (API de JavaScript para Visio)

Se aplica a: _Visio Online_

Representa la clase Page.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción|
|:---------------|:--------|:----------|
|height|entero|Devuelve el alto de la página. Solo lectura.|
|Index|int|Índice de la página. Solo lectura.|
|isBackground|bool|Si la página es una página de fondo o no. Solo lectura.|
|name|string|Nombre de la página. Solo lectura.|
|width|entero|Devuelve el ancho de la página. Solo lectura.|

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción|
|:---------------|:--------|:----------|
|comentarios|[CommentCollection](commentcollection.md)|Devuelve la colección de comentarios. Solo lectura.|
|formas|[ShapeCollection](shapecollection.md)|Formas en la página. Solo lectura.|
|vista|[PageView](pageview.md)|Devuelve la vista de la página. Solo lectura.|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[activate()](#activate)|vacío|Establece la página como la página activa del documento.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

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
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
void
