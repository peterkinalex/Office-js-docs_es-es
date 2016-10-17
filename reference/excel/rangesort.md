# <a name="rangesort-object-(javascript-api-for-excel)"></a>Objeto RangeSort (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS, Office 2016_

Administra operaciones de ordenación en objetos Range.

## <a name="properties"></a>Propiedades

Ninguno

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|void|Realiza una operación de ordenación.|

## <a name="method-details"></a>Detalles del método


### <a name="apply(fields:-sortfield[],-matchcase:-bool,-hasheaders:-bool,-orientation:-string,-method:-string)"></a>apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)
Realiza una operación de ordenación.

#### <a name="syntax"></a>Sintaxis
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|fields|SortField[]|La lista de condiciones por las que realizar la ordenación.|
|matchCase|bool|Opcional. Indica si la ordenación de cadenas distingue mayúsculas de minúsculas.|
|hasHeaders|bool|Opcional. Si el rango tiene un encabezado.|
|orientation|string|Opcional. Indica si la operación ordena filas o columnas.  Los valores posibles son: Rows, Columns|
|method|string|Opcional. Método de ordenación que se usa para los caracteres chinos.  Los valores posibles son: PinYin, StrokeCount|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```