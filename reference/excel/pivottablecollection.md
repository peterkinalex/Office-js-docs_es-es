# <a name="pivottablecollection-object-javascript-api-for-excel"></a>Objeto PivotTableCollection (API de JavaScript para Excel)

Representa una colección de todas las tablas dinámicas que forman parte del libro o de la hoja de cálculo.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|elementos|[PivotTable[]](pivottable.md)|Una colección de objetos de tabla dinámica. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|entero|Obtiene el número de tablas dinámicas de una colección.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|Obtiene una tabla dinámica por nombre.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[PivotTable](pivottable.md)|Obtiene una tabla dinámica por nombre. Si no existe la tabla dinámica, devolverá un objeto NULL.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshAll()](#refreshall)|void|Actualiza todas las tablas dinámicas de la colección.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="getcount"></a>getCount()
Obtiene el número de tablas dinámicas de una colección.

#### <a name="syntax"></a>Sintaxis
```js
pivotTableCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
entero

### <a name="getitemname-string"></a>getItem(name: string)
Obtiene una tabla dinámica por nombre.

#### <a name="syntax"></a>Sintaxis
```js
pivotTableCollectionObject.getItem(name);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Nombre de la tabla dinámica que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[PivotTable](pivottable.md)

### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
Obtiene una tabla dinámica por nombre. Si no existe la tabla dinámica, devolverá un objeto NULL.

#### <a name="syntax"></a>Sintaxis
```js
pivotTableCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|name|string|Nombre de la tabla dinámica que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[PivotTable](pivottable.md)

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
