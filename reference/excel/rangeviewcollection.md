# <a name="rangeviewcollection-object-javascript-api-for-excel"></a>Objeto RangeViewCollection (API de JavaScript para Excel)

Representa una colección de objetos de hoja de cálculo que forman parte del libro.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|items|[RangeView[]](rangeview.md)|Una colección de objetos RangeView. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[getItemAt(index: number)](#getitematindex-number)|[RangeView](rangeview.md)|Obtiene una fila RangeView mediante su índice. Indexado con cero.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy que se ha creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtiene una fila RangeView mediante su índice. Indexado con cero.

#### <a name="syntax"></a>Sintaxis
```js
rangeViewCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Índice de la fila visible.|

#### <a name="returns"></a>Valores devueltos
[RangeView](rangeview.md)

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
