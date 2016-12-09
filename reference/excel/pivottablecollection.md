# <a name="pivottablecollection-object-javascript-api-for-excel"></a>Objeto PivotTableCollection (API de JavaScript para Excel)

Representa una colección de todas las tablas dinámicas que forman parte del libro o de la hoja de cálculo.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|items|[PivotTable[]](pivottable.md)|Una colección de objetos de tabla dinámica. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|Obtiene una tabla dinámica por nombre.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(name: string)](#getitemornullname-string)|[PivotTable](pivottable.md)|Obtiene una tabla dinámica por nombre. Si la tabla dinámica no existe, la propiedad isNull del objeto devuelto será True.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy que se ha creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshAll()](#refreshall)|void|Actualiza todas las tablas dinámicas de la colección.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="getitemname-string"></a>getItem(name: string)
Obtiene una tabla dinámica por nombre.

#### <a name="syntax"></a>Sintaxis
```js
pivotTableCollectionObject.getItem(name);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Nombre de la tabla dinámica que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[PivotTable](pivottable.md)

### <a name="getitemornullname-string"></a>getItemOrNull(name: string)
Obtiene una tabla dinámica por nombre. Si la tabla dinámica no existe, la propiedad isNull del objeto devuelto será True.

#### <a name="syntax"></a>Sintaxis
```js
pivotTableCollectionObject.getItemOrNull(name);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Nombre de la tabla dinámica que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[PivotTable](pivottable.md)

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

### <a name="refreshall"></a>refreshAll()
Actualiza todas las tablas dinámicas de la colección.

#### <a name="syntax"></a>Sintaxis
```js
pivotTableCollectionObject.refreshAll();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void
