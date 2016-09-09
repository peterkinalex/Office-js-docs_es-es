# Objeto PageCollection (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una colección de páginas.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de páginas de una colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-count)|
|items|[Page[]](page.md)|Colección de objetos de página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-items)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getByTitle(title: string)](#getbytitletitle-string)|[PageCollection](pagecollection.md)|Obtiene la colección de páginas con el título especificado.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getByTitle)|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Page](page.md)|Obtiene una página por su id. o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Page](page.md)|Obtiene una página según su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-load)|

## Detalles del método


### getByTitle(title: string)
Obtiene la colección de páginas con el título especificado.

#### Sintaxis
```js
pageCollectionObject.getByTitle(title);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|title|string|El título de la página.|

#### Valores devueltos
[PageCollection](pagecollection.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // Get all the pages in the current section.
    var allPages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.
    allPages.load("id"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Get the sections with the specified name.
            var todoPages = allPages.getByTitle("Todo list");

            // Queue a command to load the section. 
            // For best performance, request specific properties.
            todoPages.load("id,title"); 

            return context.sync()
                .then(function () {

                    // Iterate through the collection or access items individually by index.
                    if (todoPages.items.length > 0) {
                        console.log("Page title: " + todoPages.items[0].title);
                        console.log("Page ID: " + todoPages.items[0].id);
                    }
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

### getItem(index: number or string)
Obtiene una página por id. o por su índice en la colección. Solo lectura.

#### Sintaxis
```js
pageCollectionObject.getItem(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number o string|El id. de la página, o bien la ubicación del índice del bloc de notas en la colección.|

#### Valores devueltos
[Page](page.md)

### getItemAt(index: number)
Obtiene una página según su posición en la colección.

#### Sintaxis
```js
pageCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[Page](page.md)

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

**Items**
```js
OneNote.run(function (context) {
    
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
    
    // Queue a command to load the id and title for each page.            
    pages.load('id,title');
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Display the properties.
            $.each(pages.items, function(index, page) {
                console.log(page.title);
                console.log(page.id);
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

