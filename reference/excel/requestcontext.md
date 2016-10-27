# <a name="requestcontext-object-(javascript-api-for-excel)"></a>Objeto RequestContext (API de JavaScript para Excel)

El objeto RequestContext facilita las solicitudes para la aplicación de Excel. Dado que el complemento de Office y la aplicación de Excel se ejecutan en dos procesos diferentes, hace falta contexto de solicitud para acceder desde el complemento a Excel y a objetos relacionados, como hojas de cálculo, tablas, etc. 

## <a name="properties"></a>Propiedades
Ninguno

## <a name="methods"></a>Métodos

| Método         | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |Rellena el objeto proxy creado en la capa de JavaScript con la propiedad y las opciones especificadas en el parámetro.|

## <a name="api-specification"></a>Especificación de API

### <a name="load(object:-object,-option:-object)"></a>load(object: object, option: object)
Rellena el objeto proxy creado en la capa de JavaScript con la propiedad y las opciones especificadas en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
requestContextObject.load(object, loadOption);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:----------------|:--------|:----------|
|object|object|Opcional. Especifique el nombre del objeto que se va a cargar.|
|option|[loadOption](loadoption.md)|Opcional. Especifique las opciones de carga, como select, expand, skip y top. Consulte el objeto loadOption para obtener más detalles.|

#### <a name="returns"></a>Valores devueltos
void

##### <a name="examples"></a>Ejemplos

En el ejemplo siguiente se cargan los valores de propiedad de un intervalo y se copian a otro intervalo.

```js
Excel.run(function (ctx) { 
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
    ctx.load(range, "values");
    return ctx.sync().then(function() {
        var myvalues=range.values;
        ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = myvalues;
        console.log(range.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
})
```