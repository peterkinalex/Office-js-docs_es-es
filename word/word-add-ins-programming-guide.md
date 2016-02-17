# Introducción a la programación de complementos de Word

_Se aplica a: Word 2016, Word para iPad, Word para Mac_

Word 2016 presenta un nuevo modelo de objetos para trabajar con objetos de Word. Este modelo de objetos es una adición al modelo de objetos existente proporcionado por Office.js para crear complementos de Word. El acceso a este modelo de objetos se efectúa a través de JavaScript alojado en una aplicación web.

## manifiesto

La nueva API de JavaScript del complemento de Word usa el mismo formato de manifiesto que se emplea en el modelo de complemento de Office 2013. El manifiesto describe dónde se hospeda el complemento, cómo se muestra, los permisos y otra información. Obtenga más información sobre cómo personalizar los [manifiestos de complemento](https://msdn.microsoft.com/en-us/library/office/fp161044.aspx). 

Dispone de varias opciones para publicar manifiestos de complemento de Word. Infórmese sobre cómo puede [publicar su complemento de Office](https://msdn.microsoft.com/EN-US/library/office/fp123515.aspx) en un recurso compartido de red, un catálogo de aplicaciones o la Tienda Office.

## Información sobre la API de JavaScript para Word

La API de JavaScript para Word se carga a través de Office.js. Proporciona un conjunto de objetos proxy de JavaScript que se usan para poner en cola un conjunto de comandos que interactúan con el contenido de un documento de Word. Estos comandos se ejecutan como un lote. Los resultados del lote son acciones realizadas en el documento de Word, como insertar contenido y sincronizar los objetos de Word con los objetos proxy de JavaScript. 

### Ejecutar el complemento

Veamos a continuación qué necesita para ejecutar el complemento. Todos los complementos deben tener un controlador de eventos Office.initialize.  Lea [Información sobre la API](https://msdn.microsoft.com/EN-US/library/fp160953.aspx) para obtener más información sobre la inicialización del complemento.  

El complemento de Word se ejecuta pasando una función al método Word.run(). La función que se pasa al método de ejecución debe tener un argumento de contexto. Este [objeto de contexto](word-add-ins-javascript-reference/requestcontext.md) es diferente del objeto de contexto que se obtiene del objeto de Office, aunque se usa con la misma finalidad, que consiste en interactuar con el entorno de tiempo de ejecución de Word. El objeto de contexto proporciona acceso al modelo de objetos de JavaScript de Word. Veamos ahora los comentarios y el código de un complemento básico de Word:

**Ejemplo 1. Inicialización y ejecución de un complemento de Word**

```javascript
    (function () {
        "use strict";

        // The initialize event handler is run each time the page is loaded.
        Office.initialize = function (reason) {
            
            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // Set your initialization code. You can use the reason 
                // argument to determine how the add-in was loaded.
                // You can also load saved settings from the Office object.
            });
        };

        // Run a batch operation against the Word object model.
        // Use the context argument to get access to the Word document.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;
        })
    })();
```

En el ejemplo 1 se muestra el código básico necesario para crear un complemento de Word. Inicializa Office.js y contiene un método de ejecución para interactuar con el documento de Word.

### Objetos proxy

El modelo de objetos de JavaScript de Word se acopla libremente a los objetos de Word. Los objetos de JavaScript de Word son objetos proxy para los objetos reales de un documento de Word. Las acciones llevadas a cabo en los objetos proxy no se realizan en Word y el estado del documento de Word no se realiza en los objetos proxy mientras no se sincronice el estado del documento. El estado del documento se sincroniza cuando se ejecuta context.sync(). Básicamente, el método sync() ejecuta el conjunto de comandos que se encuentran en la cola para cada objeto proxy.  En el ejemplo 2 se muestra la creación de un objeto proxy de cuerpo y un comando en la cola para cargar la propiedad Text en el objeto proxy de cuerpo y, a continuación, la sincronización del cuerpo del documento de Word con el objeto proxy de cuerpo. 

**Ejemplo 2. Sincronización del cuerpo del documento con el objeto proxy de cuerpo**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        // The body object hasn't been set with any property values. 
        var body = context.document.body;

        // Queue a command to load the text property for the proxy document body object.
        context.load(body, 'text');

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```

### Cola de comandos

Los objetos proxy de Word disponen de métodos para acceder y actualizar el modelo de objetos. Estos métodos se ejecutan secuencialmente en el orden en el que se pusieron en la cola en el lote. Antes de que se realice una llamada a context.sync(), se forma un lote de comandos. Se ejecutarán todos los comandos puestos en la cola en todos los objetos que usan el contexto.  

En el ejemplo 3, mostraremos cómo funciona la cola de comandos. Cuando se llama a context.sync(), lo primero que ocurre es que se ejecuta en Word el [comando que va a cargar](Word%20Add-ins%20JavaScript%20Reference/loadoption.md) el texto del cuerpo. A continuación, se ejecuta el comando que va a insertar texto en el cuerpo en Word. Después, los resultados se devuelven al objeto proxy de cuerpo. El valor de la propiedad body.text de JavaScript de Word será el valor del cuerpo del documento de Word <u>antes</u> de que el texto se insertase en el documento de Word. 

**Ejemplo 3. Ejecutar un lote de comandos.**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text in the proxy body object.
        context.load(body, 'text');

        // Queue a command to insert text into the end of the Word document body.
        body.insertText('This is text inserted after loading the body.text property',
                        Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```

## Denos su opinión

Su opinión es importante para nosotros. 

* Consulte los documentos y háganos saber todas las preguntas y las dificultades que le planteen [enviando un problema](https://github.com/OfficeDev/office-js-docs/issues) directamente en este repositorio.
* Infórmenos sobre su experiencia de programación, lo que le gustaría ver en versiones futuras, ejemplos de código, etc. Use [este sitio](http://officespdev.uservoice.com/) para enviar sus sugerencias e ideas.


## Recursos adicionales

* [Complementos de Word](word-add-ins.md)
* [Referencia de JavaScript de complementos de Word](word-add-ins-javascript-reference.md)
* [Complementos de Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Introducción a los complementos de Office](http://dev.office.com/getting-started/addins)
* &lt;a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=Word"&gt;Complementos de Word en GitHub&lt;/a&gt;
* [Explorador de fragmentos de código para Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)

