
# Agregar y quitar datos adjuntos de un elemento en un formulario de redacción en Outlook

Puede usar los métodos [addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) y [addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) para adjuntar un archivo y un elemento de Outlook respectivamente al elemento que el usuario está redactando. Ambos son métodos asincrónicos, lo que significa que la ejecución puede continuar sin necesidad de esperar a que se complete la acción de agregar los datos adjuntos. Según la ubicación original y el tamaño de los datos adjuntos que se agregan, la llamada asincrónica para agregar datos adjuntos puede tardar un tiempo en completarse. Si existen tareas que dependen de que finalice la acción, debería llevarlas a cabo con un método de devolución de llamada. Este método de devolución de llamada es opcional y se invoca cuando se completa la carga de los datos adjuntos. El método de devolución de llamada usa un objeto [AsyncResult](http://dev.outlook.com/reference/add-ins/simple-types.md) como parámetro de salida que proporciona los estados, errores y valores devueltos de la acción de agregar datos adjuntos. Si la devolución de llamada necesita otros parámetros, puede especificarlos en el parámetro opcional _options.aysncContext_.  _options.asyncContext_ puede ser de cualquiera de los tipos que el método de devolución de llamada espera.

Por ejemplo, puede definir _options.asyncContext_ como un objeto JSON que contenga uno o varios pares clave-valor con el carácter ":" como separador entre una clave y el valor, y el carácter "," como separador entre pares clave-valor. Encontrará más ejemplos de cómo [pasar parámetros opcionales a métodos asincrónicos](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) en la plataforma de complementos de Office, en [Programación asincrónica en complementos de Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md). En el ejemplo siguiente se muestra cómo usar el parámetro **asyncContext** para pasar dos argumentos a un método de devolución de llamada:




```js
{ asyncContext: { var1: 1, var2: 2} }
```

Puede comprobar si la llamada del método asincrónico se completó correctamente o no con las propiedades  **status** y **error** del objeto **AsyncResult**. Si la operación de agregar datos adjuntos se completa correctamente, puede usar la propiedad  **AsyncResult.value** para obtener el identificador de los datos adjuntos. Este identificador es un número entero que podrá usar posteriormente para quitar los datos adjuntos.


 >**Nota**  Solo se recomienda usar el id. de datos adjuntos para quitar datos adjuntos si el mismo complemento los agregó en la misma sesión. En Outlook Web App y OWA para dispositivos, el id. de datos adjuntos solo es válido en la misma sesión. Una sesión finaliza cuando el usuario cierra el complemento o si empieza a redactar un formulario en línea y posteriormente aparece el formulario en línea como elemento emergente en otra ventana.


## Adjuntar un archivo

Para adjuntar un archivo a un mensaje o una cita en un formulario de redacción, use el método  **addFileAttachmentAsync** y especifique el URI del archivo. Si el archivo está protegido, puede incluir un token de autenticación o identidad apropiado como parámetro de cadena de consulta de URI. Exchange realizará una llamada al URI para obtener los datos adjuntos y el servicio web que protege el archivo tendrá que usar el token como forma de autenticación.

El ejemplo siguiente de JavaScript es un complemento de redacción que adjunta un archivo, picture.png, desde un servidor web, al mensaje o la cita que se está redactando. El método de devolución de llamada usa  **asyncResult** como parámetro, comprueba el estado de los datos adjuntos y obtiene el id. de datos adjuntos si la operación de adjuntar se completa correctamente.




```js
var mailbox;
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

Office.initialize = function () {
    mailbox = Office.context.mailbox;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID. 
        // You can optionally pass any object that you would  
        // access in the callback method as an argument to  
        // the asyncContext parameter.
        mailbox.item.addFileAttachmentAsync(
            attachmentURI,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Get the ID of the attached file.
                    attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Adjuntar un elemento de Outlook

Para adjuntar un elemento de Outlook (por ejemplo, un mensaje de correo electrónico, un calendario o un contacto) a un mensaje o una cita en un formulario de redacción, especifique el identificador de los servicios Web Exchange (EWS) del elemento y use el método  **addItemAttachmentAsync**. Para obtener el identificador de EWS de un elemento de correo electrónico, calendario, contacto o tarea en el buzón del usuario, use el método [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) y obtenga acceso a la operación de EWS [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx). La propiedad [item.itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.md) también proporciona el identificador de EWS de un elemento existente en un formulario de lectura.

La siguiente función de JavaScript,  `addItemAttachment`, amplía el primer ejemplo anterior y agrega un elemento como datos adjuntos al correo electrónico o la cita que se está redactando. La función usa como argumento el id. de EWS del elemento que se está adjuntando. Si la operación de adjuntar se completa correctamente, se obtiene el id. de datos adjuntos para futuras tareas de procesamiento, incluida la acción de quitar dichos datos adjuntos en la misma sesión.




```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(ID) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.addItemAttachmentAsync(
        ID,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```


 >**Nota**  Puede usar un complemento de redacción para adjuntar una instancia de una cita periódica en Outlook Web App o OWA para dispositivos. Sin embargo, en un cliente enriquecido de Outlook, al intentar adjuntar una instancia se adjuntarán las series periódicas (la cita principal).


## Quitar datos adjuntos


Para quitar datos adjuntos de un elemento o archivo de un mensaje o una cita en un formulario de redacción, especifique el identificador del dato adjunto correspondiente y use el método [removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md). Debe quitar solamente los datos adjuntos que el complemento agregó en la misma sesión. Asegúrese de que el identificador del dato adjunto corresponda a datos adjuntos válidos o el método devolverá un error. Al igual que los métodos  **addFileAttachmentAsync** y **addItemAttachmentAsync**,  **removeAttachmentAsync** es un método asincrónico. Debe proporcionar un método de devolución de llamada para comprobar el estado y los posibles errores mediante el objeto de parámetro de salida **AsyncResult**. También puede transferir otros parámetros al método de devolución de llamada con el parámetro opcional  **asyncContext**,que es un objeto JSON de pares clave-valor.

La siguiente función de JavaScript,  `removeAttachment`, sigue ampliando los ejemplos anteriores y quita los datos adjuntos especificados de un correo electrónico o una cita que se está redactando. La función usa como argumento el id. de los datos adjuntos que se van a quitar. Puede obtener el id. de los datos adjuntos después de una llamada al método  **addFileAttachmentAsync** o **addItemAttachmentAsync** realizada correctamente y almacenarlo para una llamada posterior al método **removeAttachmentAsync**.




```js
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be 
// removed. 
function removeAttachment(ID) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.removeAttachmentAsync(
        ID,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```


## Consejos para agregar y quitar datos adjuntos


Si el complemento de redacción agrega y quita datos adjuntos, estructure el código de forma que transfiera un id. de datos adjuntos válido a la llamada para quitar datos adjuntos, y administre el caso cuando  **AsyncResult.error** devuelva **InvalidAttachmentId**. Según la ubicación y el tamaño de los datos adjuntos, la acción de adjuntar un archivo o elemento puede tardar un tiempo. El ejemplo siguiente contiene una llamada a  **addFileAttachmentAsync**,  `write` y **removeAttachmentAsync**. Lo lógico es pensar que las llamadas se ejecutarán de forma secuencial una tras otra.


```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

// Gets the current time in minutes, seconds and milliseconds.
function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);
            }
            write ('(3): ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
    'attachmentID is: ' + attachmentID);

Office.context.mailbox.item.removeAttachmentAsync(
        attachmentID,      
        { asyncContext: null },
       function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(5): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {           
                write('(6): ' + minutesSecondsMilliSeconds() + ' ' + 
                    ID of removed attachment: ' + asyncResult.value);
            }
        });


```

Aunque  **addFileAttachmentAsync** se inicia antes de **removeAttachmentAsync**, debido a que  **addFileAttachmentAsync** es asincrónico, las llamadas a `write` y **removeAttachmentAsync** pueden comenzar antes de que finalice **addFileAttachmentAsync**. Cuando esto ocurre,  `attachmentID` permanece como **undefined** y se obtiene un error para la llamada a **removeAttachmentAsync**, como en el ejemplo siguiente:




```
 (4): 46:18:245 attachmentID is: undefined
Error executing code: Sys.ArgumentException: Sys.ArgumentException: Value does not fall within the expected range. Parameter name: attachmentId
 (2): 46:18:255 ID of added attachment: 0
 (3): 46:18:262 Finishing addFileAttachmentAsync callback method.
```

Una forma de evitar esto es comprobar que  `attachmentID` esté definido antes de llamar a **removeAttachmentAsync**. Otra forma es iniciar la llamada a  **removeAttachmentAsync** desde el método de devolución de llamada de **addFileAttachmentAsync**, tal como se muestra en el ejemplo siguiente:




```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1) ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2) ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);

                // Move the write and removeAttachmentAsync calls here 
                // inside the addFileAttachmentAsync callback, after the 
                // attaching has succeeded.
                write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'attachmentID is: ' + attachmentID);

                Office.context.mailbox.item.removeAttachmentAsync(
                    attachmentID,
                    { asyncContext: null },
                    function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed){
                            write('(5) ' + minutesSecondsMilliSeconds() + ' ' + 
                                asyncResult.error.message);
                        }
                        else {
                            write('(6) ' + minutesSecondsMilliSeconds() + ' ' + 
                                'ID of removed attachment: ' + attachmentID);
                        }
                    });
            }

            write('(3) ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

A continuación verá un ejemplo del resultado:




```
(2) 49:25:775 ID of added attachment: 1
(4) 49:25:782 attachmentID is: 1
(3) 49:25:783 Finishing addFileAttachmentAsync callback method.
(6) 49:25:789 ID of removed attachment: 1
```

Tenga en cuenta que la devolución de llamada para  **removeAttachmentAsync** se anida dentro de la devolución de llamada para **addFileAttachmentAsync**. Debido a que  **addFileAttachmentAsync** y **removeAttachmentAsync** son asincrónicos, la última línea en la devolución de llamada para **addFileAttachmentAsync** se puede ejecutar antes de que finalice la devolución de llamada para **removeAttachmentAsync**.


## Recursos adicionales



- [Crear complementos de Outlook para formularios de redacción](../outlook/compose-scenario.md)
    
- [Programación asíncrona en los complementos de Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    


