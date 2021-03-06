# <a name="requestcontext-object-(javascript-api-for-word)"></a>Objeto RequestContext (API de JavaScript para Word)

El objeto RequestContext facilita las solicitudes a la aplicación de Word desde el complemento de Word, ya que las dos aplicaciones se ejecutan en procesos diferentes.

_Se aplica a: Word 2016, Word para iPad, Word para Mac, Word Online_

## <a name="properties"></a>Propiedades
Ninguno

## <a name="methods"></a>Métodos

| Método         | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |Rellena el objeto proxy creado en la capa de JavaScript con la propiedad y las opciones especificadas en el parámetro.|
|[sync()](#sync)  |Objeto Promise |Envía la cola de solicitudes a Word y devuelve un objeto Promise, que puede usarse para encadenar más acciones.|

## <a name="method-details"></a>Detalles del método

### <a name="load(object:-object,-option:-object)"></a>load(object: object, option: object)
Rellena el objeto proxy creado en la capa de JavaScript con la propiedad y las opciones especificadas en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
requestContextObject.load(object, loadOption);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:----------------|:--------|:----------|
|object|object|Opcional. Especifique el nombre del objeto que se va a cargar.|
|option|[loadOption](loadoption.md)|Opcional, pero se usa como procedimiento recomendado. Especifique las opciones de carga, como select, expand, skip y top. |

#### <a name="returns"></a>Valores devueltos
void

##### <a name="examples"></a>Ejemplos

En el ejemplo siguiente se muestra cómo se usa el contexto de solicitud para cargar la propiedad Text en una colección de párrafo.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the text property for all of the paragraphs.
    context.load(paragraphs, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a a set of commands to get the HTML of the first paragraph.
        var html = paragraphs.items[0].getHtml();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

#### <a name="additional-information"></a>Información adicional

Es necesario llamar a load() después de agregar objetos de los que se realiza seguimiento.

### <a name="sync()"></a>sync()
Envía la cola de solicitudes a Word y devuelve un objeto Promise, que puede usarse para encadenar más acciones.

#### <a name="syntax"></a>Sintaxis
```js
requestContextObject.sync();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
Objeto Promise.

#### <a name="examples"></a>Ejemplos

En el ejemplo siguiente se muestra el método de sincronización usado dos veces: 1) cargar la colección de controles de contenido con la propiedad Text de cada control de contenido y 2) borrar el contenido del primer control de contenido de la colección.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;

    // Queue a command to load the content controls collection.
    contentControls.load('text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {

            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });
        }

    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

## <a name="support-details"></a>Detalles de compatibilidad
Use el [conjunto de requisitos](../office-add-in-requirement-sets.md) en las comprobaciones en tiempo de ejecución para asegurarse de que la aplicación es compatible con la versión de host de Word. Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).
