# Objeto RangeSort (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS y Office 2016_

Administra operaciones de ordenación en objetos Range.

## Propiedades

Ninguno

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|void|Realiza una operación de ordenación.|

## Detalles del método


### apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)
Realiza una operación de ordenación.

#### Sintaxis
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|fields|SortField[]|La lista de condiciones por las que realizar la ordenación.|
|matchCase|bool|Opcional. Indica si la ordenación de cadenas distingue mayúsculas de minúsculas.|
|hasHeaders|bool|Opcional. Si el rango tiene un encabezado.|
|orientation|string|Opcional. Indica si la operación ordena filas o columnas.  Los valores posibles son: Rows, Columns|
|method|string|Opcional. Método de ordenación que se usa para los caracteres chinos.  Los valores posibles son: PinYin, StrokeCount|

#### Valores devueltos
void

#### Ejemplos
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