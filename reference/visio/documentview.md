# <a name="documentview-object-javascript-api-for-visio"></a>Objeto DocumentView (API de JavaScript para Visio)

Se aplica a: _Visio Online_

Representa la clase DocumentView.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción|
|:---------------|:--------|:----------|
|disableHyperlinks|bool|Deshabilita los hipervínculos.|
|disablePan|bool|Deshabilita la panorámica.|
|disableZoom|bool|Deshabilita el zoom.|
|hideDiagramBoundry|bool|Oculta el límite del diagrama.|

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## <a name="method-details"></a>Detalles del método


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
