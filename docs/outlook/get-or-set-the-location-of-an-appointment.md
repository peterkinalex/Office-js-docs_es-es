
# Obtener o definir la ubicación al redactar una cita en Outlook

La API de JavaScript para Office proporciona métodos asincrónicos ([getAsync](../../reference/outlook/Location.md) y [setAsync](../../reference/outlook/Location.md)) para obtener y establecer la ubicación de una cita que el usuario está redactando. Estos métodos asincrónicos solo están disponibles para complementos de redacción. Para usar estos métodos, asegúrese de configurar el manifiesto del complemento correctamente para que Outlook active el complemento en los formularios de redacción, como se describe en [Crear complementos de Outlook para formularios de redacción](../outlook/compose-scenario.md).

La propiedad [location](../../reference/outlook/Office.context.mailbox.item.md) está disponible para el acceso de lectura tanto en formularios de redacción como de lectura en las citas. En un formulario de lectura, puede obtener acceso a la propiedad directamente desde el objeto principal, como en:




```js
item.location
```

Sin embargo, en un formulario de redacción, dado que tanto el usuario como su complemento pueden insertar o cambiar la ubicación al mismo tiempo, debe usar el método asincrónico  **getAsync** para obtener la ubicación, tal como se muestra a continuación:




```js
item.location.getAsync
```

La propiedad  **location** se encuentra disponible para acceso de escritura solo en los formularios de redacción de citas, pero no en los formularios de lectura.

Como ocurre con la mayoría de métodos asincrónicos de la API de JavaScript para Office, **getAsync** y **setAsync** usan parámetros de entrada opcionales. Para más información sobre cómo especificar estos parámetros, vea el tema sobre cómo pasar parámetros opcionales a métodos asincrónicos en [Programación asincrónica en complementos de Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## Para obtener la ubicación


En esta sección se muestra un ejemplo de código que obtiene la ubicación de la cita que el usuario está redactando y muestra la ubicación. En este ejemplo de código se asume que hay una regla en el manifiesto del complemento que activa el complemento en un formulario de redacción para una cita, tal como se muestra a continuación.


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

Para usar  **item.location.getAsync**, proporcione un método de devolución de llamada que compruebe el estado y el resultado de la llamada asincrónica. Puede proporcionar los argumentos necesarios al método de devolución de llamada a través del parámetro opcional  _asyncContext_. Puede obtener el estado, los resultados y los errores que haya con el parámetro de salida  _asyncResult_ de la devolución de llamada. Si la llamada asincrónica se realiza correctamente, puede obtener la ubicación como una cadena con la propiedad [AsyncResult.value](../../reference/outlook/simple-types.md).




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Para establecer la ubicación


En esta sección se muestra un ejemplo de código que establece la ubicación de la cita que el usuario está redactando. De forma similar al ejemplo anterior, en este ejemplo de código se asume que hay una regla en el manifiesto del complemento que activa el complemento en un formulario de redacción para una cita.

Para usar  **item.location.setAsync**, especifique una cadena de hasta 255 caracteres en el parámetro de datos. Opcionalmente, puede proporcionar un método de devolución de llamada y cualquier argumento para el método de devolución de llamada en el parámetro  _asyncContext_. Debe comprobar el estado, el resultado y los mensajes de error en el parámetro de salida de  _asyncResult_ de la devolución de llamada. Si la llamada asincrónica se realiza correctamente, **setAsync** inserta la cadena de ubicación especificada como texto sin formato y sobrescribe la ubicación existente.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Recursos adicionales



- [Obtener y definir datos de elementos en un formulario de redacción de Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Obtención y definición de datos de elementos de Outlook en los formularios de lectura o redacción](../outlook/item-data.md)
    
- [Crear complementos de Outlook para formularios de redacción](../outlook/compose-scenario.md)
    
- [Programación asíncrona en los complementos de Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Obtener, establecer o agregar destinatarios al redactar una cita o un mensaje en Outlook](../outlook/get-set-or-add-recipients.md)
    
- [Obtener o establecer el asunto al redactar una cita o un mensaje en Outlook](../outlook/get-or-set-the-subject.md)
    
- [Introducir datos en el cuerpo al redactar una cita o un mensaje en Outlook](../outlook/insert-data-in-the-body.md)
    
- [Obtener o establecer la hora al redactar una cita en Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
