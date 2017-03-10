# <a name="pageloadcompleteeventargs-object-javascript-api-for-visio"></a>Objeto PageLoadCompleteEventArgs (API de JavaScript para Visio)

Se aplica a: _Visio Online_

Proporciona información sobre la página que ha generado el evento PageLoadComplete.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción
|:---------------|:--------|:----------|
|pageName|string|Obtiene el nombre de la página que ha generado el evento PageLoad.|
|correcto|bool|Obtiene el estado de corrección o error del evento PageLoadComplete.|

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
             eventResult1 = document1.onPageLoadComplete.add(
            function (args){
                   console.log("Page name: "+args.pageName);
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
