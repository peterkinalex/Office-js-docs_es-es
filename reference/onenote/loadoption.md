# Opciones de carga de objetos 

Representa un objeto que se puede pasar al método de carga para especificar el conjunto de propiedades y relaciones que se cargarán al ejecutar el método sync() que sincroniza los estados entre objetos de OneNote y los objetos de proxy de JavaScript correspondientes en el complemento. Esto acepta las opciones como parámetros de selección y expansión para especificar el conjunto de propiedades que se cargarán en el objeto y también para permitir la paginación en la colección.

También se puede suministrar una cadena que contenga las propiedades de relaciones que se cargarán, o bien proporcionar una matriz que contenga la lista de propiedades y relaciones que se cargarán. Vea ejemplo siguiente.

```js   
object.load('<var1>,<relationship1/var2>');

// Pass the parameter as an array.
object.load(["var1", "relationship1/var2"]);
```

## Propiedades
| Propiedad     | Tipo   |Descripción|
|:---------------|:--------|:----------|
|select|object|Proporciona una lista delimitada por comas o una matriz de nombres de parámetros/relaciones que se cargarán al realizar una llamada de sincronización, como "propiedad1, relación1", [ "propiedad1", "relación1"]. Opcional.|
|expand|object|Proporciona una lista delimitada por comas o una matriz de nombres de relaciones que se cargarán al realizar una llamada de sincronización, como "relación1, relación2", [ "relación1", "relación2"]. Opcional.|
|top|int|Especifica el número de elementos de la colección consultada que se deben incluir en el resultado. Opcional.|
|skip|entero|Especifica el número de elementos de la colección que se deben omitir y no se incluyen en el resultado. Si se especifica `top`, la selección de resultados empezará después de omitir el número especificado de elementos. Opcional.|

#### Ejemplos

En el ejemplo, se obtiene el título de página y el nivel de sangría de las primeras cinco páginas en la sección actual.

```js
OneNote.run(function (context) { 
    
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages.           
    pages.load({ "select":"title,pageLevel", "top":5, "skip":0 });
    return context.sync()
        .then(function() {
            
            // Iterate through the collection of pages.    
            $.each(pages.items, function(index, page) {
                
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Indentation level: " + page.pageLevel);
                
            });
        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        })
    });
```
