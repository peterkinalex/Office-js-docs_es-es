# <a name="tablesort-object-javascript-api-for-excel"></a>Objeto TableSort (API de JavaScript para Excel)

Administra operaciones de ordenación en objetos Table.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|matchCase|bool|Indica si última ordenación de la tabla distinguía mayúsculas de minúsculas. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|method|string|Representa el método de ordenación de caracteres chinos usado por última vez para ordenar la tabla. Solo lectura. Los valores posibles son: PinYin, StrokeCount.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|fields|[SortField](sortfield.md)|Representa las condiciones actuales que se usaron por última vez para ordenar la tabla. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[apply(fields: SortField[], matchCase: bool, method: string)](#applyfields-sortfield-matchcase-bool-method-string)|void|Realiza una operación de ordenación.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[clear()](#clear)|void|Borra la ordenación que se aplica actualmente en la tabla. Aunque esto no modifica la ordenación de la tabla, borra el estado de los botones de encabezado.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy que se ha creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[reapply()](#reapply)|void|Vuelve a aplicar los parámetros de ordenación actuales a la tabla.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="applyfields-sortfield-matchcase-bool-method-string"></a>apply(fields: SortField[], matchCase: bool, method: string)
Realiza una operación de ordenación.

#### <a name="syntax"></a>Sintaxis
```js
tableSortObject.apply(fields, matchCase, method);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|fields|SortField[]|La lista de condiciones por las que realizar la ordenación.|
|matchCase|bool|Opcional. Indica si la ordenación de cadenas distingue mayúsculas de minúsculas.|
|method|string|Opcional. Método de ordenación que se usa para los caracteres chinos.  Los valores posibles son: PinYin, StrokeCount|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.apply([ 
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

### <a name="clear"></a>clear()
Borra la ordenación que se aplica actualmente en la tabla. Aunque esto no modifica la ordenación de la tabla, borra el estado de los botones de encabezado.

#### <a name="syntax"></a>Sintaxis
```js
tableSortObject.clear();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

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

### <a name="reapply"></a>reapply()
Vuelve a aplicar los parámetros de ordenación actuales a la tabla.

#### <a name="syntax"></a>Sintaxis
```js
tableSortObject.reapply();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void
