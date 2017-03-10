# <a name="boundingbox-object-javascript-api-for-visio"></a>Objeto BoundingBox (API de JavaScript para Visio)

Se aplica a: _Visio Online_

Representa el cuadro de límite de la forma.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción|
|:---------------|:--------|:----------|
|height|entero|Distancia entre los bordes superior e inferior del cuadro de límite de la forma, sin incluir los gráficos de datos asociados a la forma.|
|width|entero|Distancia entre los bordes izquierdo y derecho del cuadro de límite de la forma, sin incluir los gráficos de datos asociados a la forma.|
|x|entero|Entero que especifica la coordenada x del cuadro de límite.|
|y|entero|Entero que especifica la coordenada y del cuadro de límite.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

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
