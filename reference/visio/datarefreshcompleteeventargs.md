# <a name="datarefreshcompleteeventargs-object-javascript-api-for-visio"></a>Objeto DataRefreshCompleteEventArgs (API de JavaScript para Visio)

Se aplica a: _Visio Online_

Proporciona información sobre el documento que ha generado el evento DataRefreshComplete.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción
|:---------------|:--------|:----------|
|correcto|bool|Obtiene el estado de corrección o error del evento DataRefreshComplete.|
|documento|[Document](document.md)|Obtiene el objeto de documento que ha generado el evento DataRefreshComplete.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno

## <a name="methods"></a>Métodos
Ninguno

### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
```js
Visio.run(function (ctx) { 
  var document1= ctx.document;
               var page = document1.getActivePage();
         eventResult1 = document1.onDataRefreshComplete.add(
    function (args){
           console.log("Data Refresh Result: "+args.success);
    });

    return ctx.sync().then(function () {
       console.log("Success");
    });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
    console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
