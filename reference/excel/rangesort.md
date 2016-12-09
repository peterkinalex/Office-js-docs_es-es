# <a name="rangesort-object-javascript-api-for-excel"></a>Objeto RangeSort (API de JavaScript para Excel)

Administra operaciones de ordenación en objetos Range.

## <a name="properties"></a>Propiedades

Ninguno

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|void|Realiza una operación de ordenación.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string"></a>apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)
Realiza una operación de ordenación.

#### <a name="syntax"></a>Sintaxis
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|fields|SortField[]|La lista de condiciones por las que realizar la ordenación.|
|matchCase|bool|Opcional. Indica si la ordenación de cadenas distingue mayúsculas de minúsculas.|
|hasHeaders|bool|Opcional. Si el rango tiene un encabezado.|
|orientation|string|Opcional. Indica si la operación ordena filas o columnas.  Los valores posibles son: Rows, Columns|
|method|string|Opcional. Método de ordenación que se usa para los caracteres chinos.  Los valores posibles son: PinYin, StrokeCount|

#### <a name="returns"></a>Valores devueltos
void
