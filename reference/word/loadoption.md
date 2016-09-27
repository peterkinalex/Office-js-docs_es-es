# Objeto LoadOption (API de JavaScript para Word)

Objeto que especifica información de paginación y las propiedades que se van a cargar cuando se llama a context.sync().

_Se aplica a: Word 2016, Word para iPad, Word para Mac_

## Properties
| Propiedad     | Tipo   |Descripción|
|:---------------|:--------|:----------|
|select|object|Contiene una lista delimitada por comas o una matriz de nombres de parámetro/relación. Opcional.|
|expand|object|Contiene una lista delimitada por comas o una matriz de nombres de relación. Opcional.|
|top|int| Especifica el número máximo de elementos de colección que pueden incluirse en el resultado. Opcional. Solo se puede usar esta opción al usar la opción de notación de objetos.|
|skip|int|Especifica el número de elementos de la colección que se deben omitir y no se incluyen en el resultado. Si se especifica `top`, el conjunto de resultados empezará después de omitir el número especificado de elementos. Opcional. Solo se puede usar esta opción al usar la opción de notación de objetos.|

## Más información

Es el método preferido para especificar las propiedades e información de paginación usando un literal de cadena. Los dos primeros ejemplos muestran la forma preferida para solicitar las propiedades de tamaño de fuente y texto de los párrafos en una colección de párrafo:

<code>context.load(paragraphs, 'text, font/size');</code>

<code>paragraphs.load('text, font/size');</code>

Aquí hay un ejemplo parecido que usa la notación de objetos (incluye la paginación):

<code>context.load(paragraphs, {select: 'text, font/size',
                                expand: 'font',
                                top: 50,
                                skip: 0});</code>

<code>paragraphs.load({select: 'text, font/size',
                       expand: 'font',
                       top: 50,
                       skip: 0});</code>

Tenga en cuenta que si no especificamos las propiedades específicas del objeto de fuente en la instrucción Select, la instrucción Expand indicaría que están cargadas todas las propiedades de fuente.

## Ejemplos

En este ejemplo se muestra cómo obtener los 50 párrafos superiores del documento de Word junto con sus propiedades de tamaño de fuente y texto.

```js
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the paragraphs collection.
            var paragraphs = context.document.body.paragraphs;

            // Queue a commmand to load the text and font properties.
            // It is best practice to always specify the property set. Otherwise, all properties are
            // returned in on the object.
            context.load(paragraphs, 'text, font/size');

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

            // Insert code that works with the paragraphs loaded by context.load().
           })
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });

```

## Detalles de compatibilidad
Use el [conjunto de requisitos](../office-add-in-requirement-sets.md) en las comprobaciones en tiempo de ejecución para asegurarse de que la aplicación es compatible con la versión de host de Word. Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).
