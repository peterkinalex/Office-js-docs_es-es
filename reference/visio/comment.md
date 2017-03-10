# <a name="comment-object-javascript-api-for-visio"></a>Objeto Comment (API de JavaScript para Visio)

Se aplica a: _Visio Online_

Representa el comentario.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción
|:---------------|:--------|:----------|
|autor|string|Cadena que especifica el nombre del autor del comentario.|
|text|string|Cadena que contiene el texto del comentario.|
|date|string|Cadena que especifica la fecha de creación del comentario.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## <a name="method-details"></a>Detalles del método


### <a name="loadparam-object"></a>load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
object.load(param);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
void
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
```js
 Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Position Belt.41";
    var shape = activePage.shapes.getItem(shapeName);
    var shapecomments= shape.comments;
        shapecomments.load();
        return ctx.sync().then(function () {
             for(var i=0; i<shapecomments.items.length;i++)
        {
                    var comment= shapecomments.items[i];
            console.log("comment Author: " + comment.author);
            console.log("Comment Text: " + comment.text);
            console.log("Date " + comment.date);
        }
     });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
