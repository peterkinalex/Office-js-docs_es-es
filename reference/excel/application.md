# <a name="application-object-javascript-api-for-excel"></a>Objeto Application (API de JavaScript para Excel)

Representa la aplicación de Excel que administra el libro.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción|Conjunto req.|
|:---------------|:--------|:----------|:----------|
|calculationMode|string|Devuelve el modo de cálculo usado en el libro. Solo lectura. Los valores posibles son: `Automatic` Excel controla el recálculo; `AutomaticExceptTables` Excel controla el recálculo pero omite los cambios de las tablas; `Manual` el cálculo se realiza cuando el usuario lo solicita.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|Conjunto req.|
|:---------------|:--------|:----------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Recalcula todos los libros abiertos actualmente en Excel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy que se ha creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="calculatecalculationtype-string"></a>calculate(calculationType: string)
Recalcula todos los libros abiertos actualmente en Excel.

#### <a name="syntax"></a>Sintaxis
```js
applicationObject.calculate(calculationType);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|calculationType|string|Especifica el tipo de cálculo que se va a usar. Los valores posibles son: `Recalculate` Este es un recálculo suave y se usa principalmente para la compatibilidad con versiones anteriores. `Full` Recalcula todas las celdas que Excel ha marcado como modificadas, es decir, dependientes de datos cambiados o volátiles, y las celdas que se han marcado mediante programación como modificadas. `FullRebuild` Recalcula todas las celdas de todos los libros abiertos.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
Excel.run(function (ctx) { 
    ctx.workbook.application.calculate('Full');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="loadparam-object"></a>load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
object.load(param);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, acepta un objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
void
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
```js
Excel.run(function (ctx) { 
    var application = ctx.workbook.application;
    application.load('calculationMode');
    return ctx.sync().then(function() {
        console.log(application.calculationMode);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
