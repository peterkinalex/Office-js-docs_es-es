# Objeto WorksheetProtection (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS y Office 2016_

Representa la protección de un objeto de hoja.

## Propiedades

| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|protegido|booleano|Indica si la hoja de cálculo está protegida. Solo lectura.|

## Relaciones
| Relación | Tipo|Descripción|
|:---------------|:--------|:----------|
|opciones|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|Opciones de protección de la hoja. Solo lectura.|

## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto de proxy con los detalles de protección de la hoja.|
|[protect(options: WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoption)|void|Proteger una hoja de cálculo. Produce una excepción si se ha protegido la hoja de cálculo.|
|[unprotect()](#unprotect)|void|Desproteger una hoja de cálculo|

## Detalles del método


### load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### Sintaxis
```js
object.load(param);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### Valores devueltos
void

#### Ejemplos
Este ejemplo carga la información de protección de la hoja de cálculo activa.
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection status: " + worksheet.protection.protected);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### protect(options: WorksheetProtectionOptions)
Proteger una hoja de cálculo con las directivas de protección opcionales. Produce una excepción si se ha protegido la hoja de cálculo. 

Si se especifican opciones, es posible habilitar o deshabilitar directivas individuales. Si no se especifica una directiva, está habilitado de forma predeterminada. 

#### Sintaxis
```js
worksheetProtectionObject.protect(options);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|opciones|WorksheetProtectionOptions|Opcional. Opciones de protección de la hoja.|


#### Valores devueltos
void

#### Ejemplos
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
### unprotect()
Desproteger una hoja de cálculo. 

#### Sintaxis
```js
worksheetProtectionObject.unprotect();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos
```js
Excel.run(function (ctx) { 
	var sheet = ctx.workbook.worksheets.getItem("Sheet1");	
	sheet.protection.unprotect();
	return ctx.sync(); 
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
