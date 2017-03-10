# <a name="selection-object-javascript-api-for-visio"></a>Objeto Selection (API de JavaScript para Visio)

Se aplica a: _Visio Online_

Representa la selección de la página.

## <a name="properties"></a>Propiedades

Ninguno

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción|
|:---------------|:--------|:----------|
|formas|[ShapeCollection](shapecollection.md)|Obtiene las formas de la selección. Solo lectura.|

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
