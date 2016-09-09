# Objeto PageContentCollection (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa el contenido de una página, como una colección de objetos PageContent.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de contenidos de página de la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-count)|
|items|[PageContent[]](pagecontent.md)|Una colección de objetos PageContent. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-items)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[PageContent](pagecontent.md)|Obtiene un objeto PageContent por id. o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[PageContent](pagecontent.md)|Obtiene un contenido de página según su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-load)|

## Detalles del método


### getItem(index: number or string)
Obtiene un objeto PageContent por id. o por su índice en la colección. Solo lectura.

#### Sintaxis
```js
pageContentCollectionObject.getItem(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number o string|El id. del objeto PageContent, o bien la ubicación del índice del bloc de notas en la colección.|

#### Valores devueltos
[PageContent](pagecontent.md)

### getItemAt(index: number)
Obtiene un contenido de página según su posición en la colección.

#### Sintaxis
```js
pageContentCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[PageContent](pagecontent.md)

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
            console.log("The first page content item is of type: " + firstPageContent.type);
            return context.sync();
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
### Ejemplos de acceso a la propiedad

**items**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Queue a command to load the type of each pageContent.
    pageContents.load("type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            $.each(pageContents.items, function(index, pageContent) {
                console.log("PageContent type: " + pageContent.type);
            });
        });
})                
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**recorrido para esquemas**
```js
OneNote.run(function (context) {
   var page = context.application.getActivePage();
   var pageContents = page.contents;
   pageContents.load('type');
   var outlines = [];
   return context.sync()
       .then(function () {    
              $.each(pageContents.items, function (index, pageContent) {
                     console.log(pageContent.type);
                     if (pageContent.type === 'Outline') {
                           outlines.push(pageContent);
                     }
              });
              $.each(outlines, function (index, outline) {
                     outline.load("id,paragraphs,paragraphs/type");
              });
              return context.sync();
       })
       .then(function () {
              $.each(outlines, function (index, outline) {
                     console.log("An outline was found with id : " + outline.id);
              });
              return Promise.resolve(outlines);
       });
});
```

