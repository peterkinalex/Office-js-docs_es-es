# <a name="worksheetprotection-object-javascript-api-for-excel"></a>Objeto WorksheetProtection (API de JavaScript para Excel)

Representa la protección de un objeto de hoja.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|protected|bool|Indica si la hoja de cálculo está protegida. Solo lectura. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|opciones|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|Opciones de protección de la hoja. Solo lectura. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|[protect(options: WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoptions)|vacío|Protege una hoja de cálculo. Produce un error si se ha protegido la hoja de cálculo.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[unprotect()](#unprotect)|void|Desprotege una hoja de cálculo.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="protectoptions-worksheetprotectionoptions"></a>protect(options: WorksheetProtectionOptions)
Protege una hoja de cálculo. Produce un error si se ha protegido la hoja de cálculo.

#### <a name="syntax"></a>Sintaxis
```js
worksheetProtectionObject.protect(options);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|opciones|WorksheetProtectionOptions|Opcional. Opciones de protección de la hoja.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    var range = sheet.getRange("A1:B3").format.protection.locked = false;
    sheet.protection.protect({allowInsertRows:true});
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

```
### <a name="unprotect"></a>unprotect()
Desprotege una hoja de cálculo.

#### <a name="syntax"></a>Sintaxis
```js
worksheetProtectionObject.unprotect();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void
