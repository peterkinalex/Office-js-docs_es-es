# <a name="rangeviewcollection-object-javascript-api-for-excel"></a>Objeto RangeViewCollection (API de JavaScript para Excel)

Representa una colección de objetos RangeView.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|elementos|[RangeView[]](rangeview.md)|Una colección de objetos RangeView. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|entero|Obtiene el número de objetos RangeView de la colección.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[RangeView](rangeview.md)|Obtiene una fila RangeView mediante su índice. Indexado con cero.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="getcount"></a>getCount()
Obtiene el número de objetos RangeView de la colección.

#### <a name="syntax"></a>Sintaxis
```js
rangeViewCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
entero

### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtiene una fila RangeView mediante su índice. Indexado con cero.

#### <a name="syntax"></a>Sintaxis
```js
rangeViewCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|index|number|Índice de la fila visible.|

#### <a name="returns"></a>Valores devueltos
[RangeView](rangeview.md)
