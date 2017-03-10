# <a name="selectionchangedeventargs-object-javascript-api-for-visio"></a>Objeto SelectionChangedEventArgs (API de JavaScript para Visio)

Se aplica a: _Visio Online_

Proporciona información sobre la colección de formas que ha generado el evento SelectionChanged.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción
|:---------------|:--------|:----------|
|shapeNames|string[]|Obtiene la matriz de nombres de forma que ha generado el evento SelectionChanged.|
|pageName|string|Obtiene el nombre de la página que tiene el objeto de colección ShapeCollection que ha generado el evento SelectionChanged.|

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
             eventResult1 = document1.onSelectionChanged.add(
        function (args){
                   console.log("Selected Shape Name: "+args.shapeNames[0]);
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
