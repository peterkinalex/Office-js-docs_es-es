# <a name="shapemouseleaveeventargs-object-javascript-api-for-visio"></a>Objeto ShapeMouseLeaveEventArgs (API de JavaScript para Visio)

Se aplica a: _Visio Online_

Proporciona información sobre la forma que ha generado el evento MouseLeave.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción
|:---------------|:--------|:----------|
|shapeName|string|Obtiene el nombre del objeto de forma que ha generado el evento MouseLeave.|
|pageName|string|Obtiene el nombre de la página que tiene el objeto de forma que ha generado el evento MouseLeave.|

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
    eventResult2 = document1.onMouseLeave.add(
                function (args){            
                         console.log(Date.now()+":OnMouseLeave Event"+JSON.stringify(args));
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