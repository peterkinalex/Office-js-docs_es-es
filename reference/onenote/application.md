# Objeto Application (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_


Representa el objeto de nivel superior que contiene todos los objetos de OneNote a los que se puede hacer referencia globalmente, como blocs de notas, el bloc de notas activo y la sección activa.

## Propiedades

Ninguno

## Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|blocs de notas|[NotebookCollection](notebookcollection.md)|Obtiene la colección de blocs de notas que están abiertos en la instancia de la aplicación de OneNote. En OneNote Online, solo se abre un bloc de notas en la instancia de la aplicación. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-notebooks)|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getActiveNotebook()](#getactivenotebook)|[Notebook](notebook.md)|Obtiene el bloc de notas activo si existe alguno. Si no hay ningún bloc de notas activo, lanza ItemNotFound.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebook)|
|[getActiveNotebookOrNull()](#getactivenotebookornull)|[Notebook](notebook.md)|Obtiene el bloc de notas activo si existe alguno. Si no hay ningún bloc de notas activo, devuelve NULL.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebookOrNull)|
|[getActiveOutline()](#getactiveoutline)|[Outline](outline.md)|Obtiene el esquema activo si existe alguno. Si no hay ninguno, lanza ItemNotFound.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutline)|
|[getActiveOutlineOrNull()](#getactiveoutlineornull)|[Outline](outline.md)|Obtiene el esquema activo si existe. De lo contrario, devuelve NULL.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutlineOrNull)|
|[getActivePage()](#getactivepage)|[Page](page.md)|Obtiene la página activa si existe alguna. Si no hay ninguna página activa, lanza ItemNotFound.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePage)|
|[getActivePageOrNull()](#getactivepageornull)|[Page](page.md)|Obtiene la página activa si existe alguna. Si no hay ninguna página activa, devuelve NULL.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePageOrNull)|
|[getActiveSection()](#getactivesection)|[Section](section.md)|Obtiene la sección activa si existe alguna. Si no hay ninguna sección activa, lanza ItemNotFound.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSection)|
|[getActiveSectionOrNull()](#getactivesectionornull)|[Section](section.md)|Obtiene la sección activa si existe alguna. Si no hay ninguna sección activa, devuelve NULL.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSectionOrNull)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-load)|
|[navigateToPage(page: Page)](#navigatetopagepage-page)|void|Abre la página especificada en la instancia de la aplicación.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPage)|
|[navigateToPageWithClientUrl(url: string)](#navigatetopagewithclienturlurl-string)|[Page](page.md)|Obtiene la página especificada y la abre en la instancia de la aplicación.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPageWithClientUrl)|

## Detalles del método


### getActiveNotebook()
Obtiene el bloc de notas activo si existe alguno. Si no hay ningún bloc de notas activo, lanza ItemNotFound.

#### Sintaxis
```js
applicationObject.getActiveNotebook();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Notebook](notebook.md)

#### Ejemplos
```js
OneNote.run(function (context) {
        
    // Get the active notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Notebook name: " + notebook.name);
            console.log("Notebook ID: " + notebook.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### getActiveNotebookOrNull()
Obtiene el bloc de notas activo si existe alguno. Si no hay ningún bloc de notas activo, devuelve NULL.

#### Sintaxis
```js
applicationObject.getActiveNotebookOrNull();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Notebook](notebook.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // Get the active notebook.
    var notebook = context.application.getActiveNotebookOrNull();

    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // check if active notebook is set.
            if (!notebook.isNull) {
                console.log("Notebook name: " + notebook.name);
                console.log("Notebook ID: " + notebook.id);
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


### getActiveOutline()
Obtiene el esquema activo si existe alguno. Si no hay ninguno, lanza ItemNotFound.

#### Sintaxis
```js
applicationObject.getActiveOutline();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Outline](outline.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutline();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Show some properties.
            console.log("outline id: " + outline.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### getActiveOutlineOrNull()
Obtiene el esquema activo si existe. De lo contrario, devuelve NULL.

#### Sintaxis
```js
applicationObject.getActiveOutlineOrNull();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Outline](outline.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutlineOrNull();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            if (!outline.isNull) {
                console.log("outline id: " + outline.id);
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


### getActivePage()
Obtiene la página activa si existe alguna. Si no hay ninguna página activa, lanza ItemNotFound.

#### Sintaxis
```js
applicationObject.getActivePage();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Page](page.md)

#### Ejemplos
```js
OneNote.run(function (context) {
        
    // Get the active page.
    var page = context.application.getActivePage();
            
    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Page title: " + page.title);
            console.log("Page ID: " + page.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### getActivePageOrNull()
Obtiene la página activa si existe alguna. Si no hay ninguna página activa, devuelve NULL.

#### Sintaxis
```js
applicationObject.getActivePageOrNull();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Page](page.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // Get the active page.
    var page = context.application.getActivePageOrNull();

    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            if (!page.isNull) {
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Page ID: " + page.id);
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


### getActiveSection()
Obtiene la sección activa si existe alguna. Si no hay ninguna sección activa, lanza ItemNotFound.

#### Sintaxis
```js
applicationObject.getActiveSection();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Sección](section.md)

#### Ejemplos
```js
OneNote.run(function (context) {
        
    // Get the active section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Section name: " + section.name);
            console.log("Section ID: " + section.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### getActiveSectionOrNull()
Obtiene la sección activa si existe alguna. Si no hay ninguna sección activa, devuelve NULL.

#### Sintaxis
```js
applicationObject.getActiveSectionOrNull();
```

#### Parámetros
Ninguno

#### Valores devueltos
[Sección](section.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // Get the active section.
    var section = context.application.getActiveSectionOrNull();

    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if (!section.isNull) {
                // Show some properties.
                console.log("Section name: " + section.name);
                console.log("Section ID: " + section.id);
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

### navigateToPage(page: Page)
Abre la página especificada en la instancia de la aplicación.

#### Sintaxis
```js
applicationObject.navigateToPage(page);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|page|Page|La página que se abrirá.|

#### Valores devueltos
void

#### Ejemplos
```js        
OneNote.run(function (context) {
        
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // This example loads the first page in the section.
            var page = pages.items[0];
                        
            // Open the page in the application.                    
            context.application.navigateToPage(page);
                    
            // Run the queued command.
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


### navigateToPageWithClientUrl(url: string)
Obtiene la página especificada y la abre en la instancia de la aplicación.

#### Sintaxis
```js
applicationObject.navigateToPageWithClientUrl(url);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|url|string|La URL del cliente de la página que se abrirá.|

#### Valores devueltos
[Page](page.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('clientUrl');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // This example loads the first page in the section.
            var page = pages.items[0];

            // Open the page in the application.                    
            context.application.navigateToPageWithClientUrl(page.clientUrl);

            // Run the queued command.
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
