# Objeto Page (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_   


Representa una página de OneNote.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|clientUrl|string|La URL del cliente de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-clientUrl)|
|id|string|Obtiene el identificador de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-id)|
|pageLevel|int|Obtiene o establece el nivel de sangría de la página.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-pageLevel)|
|title|string|Obtiene o establece el título de la página.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-title)|
|webUrl|string|La URL de la web de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-webUrl)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|contents|[PageContentCollection](pagecontentcollection.md)|Obtiene la colección de objetos PageContent de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-contents)|
|inkAnalysisOrNull|[InkAnalysis](inkanalysis.md)|Interpretación del texto para la tinta de la página. Devuelve NULL si no hay ninguna información de análisis de tinta. Solo lectura. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-inkAnalysisOrNull)|
|parentSection|[Section](section.md)|Obtiene la sección que contiene la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-parentSection)|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[addOutline(left: double, top: double, html: String)](#addoutlineleft-double-top-double-html-string)|[Outline](outline.md)|Agrega un esquema a la página en la posición especificada.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-addOutline)|
|[copyToSection(destinationSection: Section)](#copytosectiondestinationsection-section)|[Página](page.md)|Copia esta página en la sección especificada.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-copyToSection)|
|[insertPageAsSibling(location: string, title: string)](#insertpageassiblinglocation-string-title-string)|[Page](page.md)|Inserta una nueva página antes o después de la página actual.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-insertPageAsSibling)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-load)|

## Detalles del método


### addOutline(left: double, top: double, html: String)
Agrega un Outline a la página en la posición especificada.

#### Sintaxis
```js
pageObject.addOutline(left, top, html);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|left|double|La posición izquierda de la esquina superior izquierda del Outline.|
|top|double|La posición superior de la esquina superior izquierda del Outline.|
|html|String|Una cadena HTML que describe la presentación visual del esquema. Consulte el [HTML compatible](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) para la API de JavaScript de los complementos de OneNote.|

#### Valores devueltos
[Outline](outline.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // Gets the active page.
    var page = context.application.getActivePage();

    // Queue a command to add an outline with given html. 
    var outline = page.addOutline(200, 200,
"<p>Images and a table below:</p> \
 <img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\"> \
 <img src=\"http://imagenes.es.sftcdn.net/es/scrn/6653000/6653659/microsoft-onenote-2013-01-535x535.png\"> \
 <table> \
   <tr> \
     <td>Jill</td> \
     <td>Smith</td> \
     <td>50</td> \
   </tr> \
   <tr> \
     <td>Eve</td> \
     <td>Jackson</td> \
     <td>94</td> \
   </tr> \
 </table>"     
        );

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
});
```


### copyToSection(destinationSection: Section)
Copia esta página en la sección especificada.

#### Sintaxis
```js
pageObject.copyToSection(destinationSection);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|destinationSection|Section|La sección en la que hay que copiar esta página.|

#### Valores devueltos
[Page](page.md)

#### Ejemplos
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    
    // Gets the active notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Queue a command to load sections under the notebook.
    notebook.load('sections');
    
    var newPage;
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync()
        .then(function() {
            var section = notebook.sections.items[0];
            
            // copy page to the section.
            newPage = page.copyToSection(section);
            newPage.load('id');
            return ctx.sync();
        })
        .then(function() {
            console.log(newPage.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### insertPageAsSibling(location: string, title: string)
Inserta una nueva página antes o después de la página actual.

#### Sintaxis
```js
pageObject.insertPageAsSibling(location, title);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|location|cadena|La ubicación de la nueva página relativa a la página actual.  Los valores posibles son: Before, After|
|cargo|string|El título de la nueva página.|

#### Valores devueltos
[Page](page.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var newPage = activePage.insertPageAsSibling("After", "Next Page");

    // Queue a command to load the newPage to access its data.
    context.load(newPage);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("page is created with title: " + newPage.title);
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

**contents**
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            for(var i=0; i < pageContents.items.length; i++)
            {
                var pageContent = pageContents.items[i];
                if (pageContent.type == "Outline")
                {
                    console.log("Found an outline");
                }
                else if (pageContent.type == "Image")
                {
                    console.log("Found an image");
                }
                else if (pageContent.type == "Other")
                {
                    console.log("Found a type not supported yet.");
                }
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

**webUrl**
```js
OneNote.run(function (context) {

    var app = context.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Queue a command to load the webUrl of the page.
    page.load("webUrl");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log(page.webUrl);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**inkAnalysisOrNull**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Load ink words
    page.load('inkAnalysisOrNull/paragraphs/lines/words');
    
    return ctx.sync()
        .then(function() {
            if (!page.inkAnalysisOrNull.isNull)
                console.log(page.inkAnalysisOrNull.paragraphs.length);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

