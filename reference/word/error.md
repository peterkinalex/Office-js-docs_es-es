# <a name="officeextension.error-object-(javascript-api-for-word)"></a>Objeto OfficeExtension.Error (API de JavaScript para Word)

Representa los errores que se producen al usar la API de JavaScript para Word.

_Se aplica a: Word 2016, Word para iPad, Word para Mac, Word Online_

## <a name="properties"></a>Propiedades
| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|código|cadena|Obtiene un valor que indica el tipo de error. El valor puede ser "AccessDenied", "GeneralException", "ActivityLimitReached", "InvalidArgument", "ItemNotFound" o "NotImplemented".<!-- Values come from OfficeExtension.Error and Word.ErrorCodes. -->|
|debugInfo|string|Obtiene un valor que indica lo sucedido al producirse el error. Este valor solo está diseñado para su uso durante el desarrollo y la depuración.  |
|mensaje |cadena| Obtiene una cadena legible y localizada manualmente que corresponde al código de error.|
|nombre |cadena| Obtiene un valor que siempre es "OfficeExtension.Error". |
|traceMessages |string[]| Obtiene una matriz de valores que corresponden a los mensajes de instrumentación establecidos con context.trace();. |

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|Devuelve el código de error y los valores del mensaje en el formato siguiente: "{0}: {1}", código, mensaje.|

## <a name="method-details"></a>Detalles del método

### <a name="tostring()"></a>toString()
Devuelve el código de error y los valores del mensaje en el formato siguiente: "{0}: {1}", código, mensaje.

#### <a name="syntax"></a>Sintaxis
```js
error.toString()
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
string

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert text in to the beginning of the body.
    // This will cause an OfficeExtension.Error.
    body.insertText(0);

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync();
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Error code and message: ' + error.toString());
    }
});

```

## <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

### <a name="trace-message-instrumentation"></a>Instrumentación de los mensajes de seguimiento

En el ejemplo siguiente se muestra cómo se puede instrumentar un lote de comandos para determinar dónde se ha producido un error. El primer lote inserta correctamente los dos primeros párrafos en el documento sin provocar ningún error. El segundo lote inserta correctamente el tercer y el cuarto párrafo, pero se produce un error en la llamada para insertar el quinto. El resto de comandos tras el comando que ha producido el error del lote no se ejecuta, incluido el comando que agrega el quinto mensaje de seguimiento. En este caso, el error se ha producido tras insertar el cuarto párrafo y antes de agregar el quinto mensaje de seguimiento.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    // Start a batch of commands.
    body.insertParagraph('1st paragraph', Word.InsertLocation.end);
    // Queue a command for instrumenting this part of the batch.
    context.trace('1st paragraph successful');

    body.insertParagraph('2nd paragraph', Word.InsertLocation.end);
    context.trace('2nd paragraph successful');

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Queue a commmand to insert the paragraph at the end of the document body.
        // Start a new batch of commands.
        body.insertParagraph('3rd paragraph', Word.InsertLocation.end);
        context.trace('3rd paragraph successful');

        body.insertParagraph('4th paragraph', Word.InsertLocation.end);
        context.trace('4th paragraph successful');

        // This command will cause an error. The trace messages in the queue up to
        // this point will be available via Error.traceMessages.
        body.insertParagraph(0, '5th paragraph', Word.InsertLocation.end);
        // Queue a command for instrumenting this part of the batch.
        // This trace message will not be set on Error.traceMessages.
        context.trace('5th paragraph successful');
    }).then(context.sync);
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Trace messages: ' + error.traceMessages);
    }
});

// Output: "Trace messages: 3rd paragraph successful,4th paragraph successful"

```
