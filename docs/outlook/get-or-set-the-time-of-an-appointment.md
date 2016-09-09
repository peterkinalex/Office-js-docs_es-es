
# Obtener o establecer la hora al redactar una cita en Outlook

La API de JavaScript para Office proporciona métodos asincrónicos ([Time.getAsync](../../reference/outlook/Time.md) y [Time.setAsync](../../reference/outlook/Time.md)) para obtener y establecer la hora de inicio o la hora de finalización de una cita que el usuario está redactando. Estos métodos asincrónicos solo están disponibles para complementos de redacción. Para usar estos métodos, asegúrese de configurar el manifiesto del complemento correctamente para que Outlook active el complemento en los formularios de redacción, como se describe en [Crear complementos de Outlook para formularios de redacción](../outlook/compose-scenario.md).

Las propiedades [start](../../reference/outlook/Office.context.mailbox.item.md) y [end](../../reference/outlook/Office.context.mailbox.item.md) están disponibles para las citas tanto en formularios de redacción como de lectura. En un formulario de lectura, se puede obtener acceso a las propiedades directamente desde el objeto principal, como en:




```
item.start
```

Y en:




```
item.end
```

Pero en un formulario de redacción, debido a que tanto el usuario como el complemento pueden insertar o cambiar la hora al mismo tiempo, debe usar el método asincrónico  **getAsync** para obtener la hora de inicio o finalización, tal como se muestra a continuación:




```
item.start.getAsync
```

Y en:




```
item.end.getAsync
```

Como ocurre con la mayoría de métodos asincrónicos de la API de JavaScript para Office, **getAsync** y **setAsync** usan parámetros de entrada opcionales. Para más información sobre cómo especificar estos parámetros de entrada opcionales, vea [Pasar parámetros opcionales a métodos asincrónicos](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) en [Programación asincrónica en complementos de Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## Para obtener la hora de inicio o finalización


Esta sección muestra un código de ejemplo que obtiene la hora de inicio de una cita que el usuario está redactando y muestra la hora. Puede usar el mismo código para reemplazar la propiedad  **start** por la propiedad **end** y así obtener la hora de finalización. Este código de ejemplo asume una regla en el manifiesto del complemento que activa el complemento en un formulario de redacción para una cita, tal como se muestra a continuación.


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

Para usar  **item.start.getAsync** o **item.end.getAsync**, proporcione un método de devolución de llamada que compruebe el estado y el resultado de la llamada asincrónica. Puede proporcionar los argumentos necesarios al método de devolución de llamada a través del parámetro opcional  _asyncContext_. Puede obtener el estado, los resultados y los errores que haya con el parámetro de salida  _asyncResult_ de la devolución de llamada. Si la llamada asincrónica se realiza correctamente, podrá obtener la hora de inicio como un objeto **Date** en formato UTC con la propiedad [AsyncResult.value](../../reference/outlook/simple-types.md).




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the start time of the item being composed.
        getStartTime();
    });
}

// Get the start time of the item that the user is composing.
function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Para establecer la hora de inicio o finalización


Esta sección muestra un código de ejemplo que configura la hora de inicio de una cita o mensaje que el usuario está redactando. Puede usar el mismo código y reemplazar la propiedad  **start** por la propiedad **end** para configurar la hora de finalización. Tenga en cuenta que si el formulario de redacción de la cita ya cuenta con una hora de inicio, la configuración de otra hora de inicio ajustará la hora de finalización para mantener la duración de la cita. Si el formulario de redacción de la cita ya cuenta con una hora de finalización, la configuración de otra hora de finalización ajustará tanto la duración como la hora de finalización. Si la cita se ha establecido como un evento de todo el día, la configuración de la hora de inicio ajustará la hora de finalización 24 horas más tarde y desactivará la opción de la interfaz para el evento de todo el día en el formulario de redacción.

De forma similar al ejemplo anterior, en este ejemplo de código se asume que hay una regla en el manifiesto del complemento que activa el complemento en un formulario de redacción para una cita.

Para usar  **item.start.setAsync** o **item.end.setAsync**, especifique un valor de  **Date** en formato UTC en el parámetro _dateTime_. Si obtiene una fecha según una entrada del usuario en el cliente, podrá usar [mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md) para convertir el valor en un objeto **Date** en UTC. Puede proporcionar un método de devolución de llamada opcional y los argumentos que necesite en el parámetro _asyncContext_. Debe comprobar el estado, el resultado y los posibles mensajes de error en el parámetro de salida  _asyncResult_ de la devolución de llamada. Si la llamada asincrónica se realiza correctamente, **setAsync** inserta la cadena de hora de inicio o finalización como texto sin formato y sobrescribe las horas de inicio o finalización que haya en el elemento.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    var startDate = new Date("September 27, 2012 12:30:00");
    
    item.start.setAsync(
        startDate,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the start time.
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
    
- [Obtener o definir la ubicación al redactar una cita en Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
