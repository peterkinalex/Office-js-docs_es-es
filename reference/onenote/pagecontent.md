# Objeto PageContent (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una región de nivel superior en una página que contiene tipos de contenido de nivel superior, como Outline o Image. Un objeto PageContent se puede asignar a una posición XY.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|id|cadena|Obtiene el identificador del objeto PageContent. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-id)|
|left|double|Obtiene o establece la posición izquierda (eje X) del objeto PageContent.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-left)|
|top|double|Obtiene o establece la posición superior (eje Y) del objeto PageContent.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-top)|
|type|string|Obtiene el tipo del objeto PageContent. Solo lectura. Los valores posibles son: Outline, Image, Other.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-type)|

## Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|image|[Image (Imagen)](image.md)|Obtiene el elemento Image del objeto PageContent. Produce una excepción si PageContentType no es Image. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-image)|
|ink|[FloatingInk](floatingink.md)|Obtiene la tinta del objeto PageContent. Produce una excepción si PageContentType no es Ink. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-ink)|
|outline|[Outline](outline.md)|Obtiene el elemento Outline del objeto PageContent. Produce una excepción si PageContentType no es Outline. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-outline)|
|parentPage|[Page](page.md)|Obtiene la página que contiene el objeto PageContent. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-parentPage)|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|Elimina el objeto PageContent.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-delete)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-load)|

## Detalles del método


### delete()
Elimina el objeto PageContent.

#### Sintaxis
```js
pageContentObject.delete();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos
```js
OneNote.run(function (context) {

    var page = context.application.getActivePage();
    var pageContents = page.contents;

    var firstPageContent = pageContents.getItemAt(0);
    firstPageContent.load('type');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if(firstPageContent.isNull === false) {
                firstPageContent.delete();
                return context.sync();
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### Sintaxis
```js
object.load(param);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### Valores devueltos
void
