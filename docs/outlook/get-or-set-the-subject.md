
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a>Obtener o establecer el asunto al redactar una cita o un mensaje en Outlook

La API de JavaScript para Office proporciona métodos asincrónicos ([subject.getAsync](../../reference/outlook/Subject.md) y [subject.setAsync](../../reference/outlook/Subject.md)) para obtener y establecer el asunto de una cita o mensaje que el usuario está redactando. Estos métodos asincrónicos solo están disponibles para complementos de redacción. Para usar estos métodos, asegúrese de configurar el manifiesto del complemento correctamente para que Outlook active el complemento en los formularios de redacción.

La propiedad  **subject** está disponible para el acceso de lectura tanto en formularios de redacción como de lectura de citas y mensajes. En un formulario de lectura, puede acceder a la propiedad directamente desde el objeto principal, como en:




```js
item.subject
```

Pero en un formulario de redacción, debido al hecho de que tanto el usuario como su complemento pueden insertar o cambiar el asunto al mismo tiempo, debe usar el método asincrónico  **getAsync** para obtener el asunto, tal como se muestra a continuación:




```js
item.subject.getAsync
```

La propiedad  **subject** está disponible para el acceso de escritura solamente en formularios de redacción (no en formularios de lectura).

Como ocurre con la mayoría de métodos asincrónicos de la API de JavaScript para Office, **getAsync** y **setAsync** usan parámetros de entrada opcionales. Para más información sobre cómo especificar estos parámetros de entrada opcionales, vea "Pasar parámetros opcionales a métodos asincrónicos" en [Programación asincrónica en complementos de Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## <a name="to-get-the-subject"></a>Para obtener el asunto


Esta sección muestra un código de ejemplo que obtiene y muestra el asunto de la cita o el mensaje que el usuario está redactando. Este código de ejemplo asume una regla en el manifiesto del complemento que activa el complemento en un formulario de redacción para una cita o un mensaje, tal como se muestra a continuación.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

Para usar  **item.subject.getAsync**, proporcione un método de devolución de llamada que compruebe el estado y el resultado de la llamada asincrónica. Puede proporcionar los argumentos necesarios al método de devolución de llamada a través del parámetro opcional  _asyncContext_. Puede obtener el estado, los resultados y los errores que haya con el parámetro de salida  _asyncResult_ de la devolución de llamada. Si la llamada asincrónica se realiza correctamente, podrá obtener el asunto como una cadena de texto sin formato usando la propiedad [AsyncResult.value](../../reference/outlook/simple-types.md).




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the subject of the item being composed.
        getSubject();
    });
}

// Get the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                write ('The subject is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="to-set-the-subject"></a>Para configurar el asunto


Esta sección muestra un código de ejemplo que configura el asunto de una cita o un mensaje que el usuario está redactando. Tal como ocurre en el ejemplo anterior, este código asume una regla en el manifiesto del complemento que activa el complemento en un formulario de redacción para una cita o un mensaje.

Para usar  **item.subject.setAsync**, especifique una cadena de hasta 255 caracteres en el parámetro de datos. De forma opcional, puede proporcionar un método de devolución de llamada y los argumentos que quiera para este método en el parámetro  _asyncContext_. Debe comprobar el estado, el resultado y cualquier posible mensaje de error en el parámetro de salida  _asyncResult_ de la devolución de llamada. Si la llamada asincrónica se produce correctamente, **setAsync** inserta la cadena de asunto especificada como texto sin formato y sobrescribe cualquier asunto que exista en ese elemento.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    var today = new Date();
    var subject;

    // Customize the subject with today's date.
    subject = 'Summary for ' + today.toLocaleDateString();

    item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
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


## <a name="additional-resources"></a>Recursos adicionales



- [Obtener y definir datos de elementos en un formulario de redacción de Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Obtener y establecer datos de elementos en formularios de lectura o redacción](../outlook/item-data.md)
    
- [Crear complementos de Outlook para formularios de redacción](../outlook/compose-scenario.md)
    
- [Programación asincrónica en los complementos de Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Obtener, establecer o agregar destinatarios al redactar una cita o un mensaje en Outlook](../outlook/get-set-or-add-recipients.md)
    
- [Introducir datos en el cuerpo al redactar una cita o un mensaje en Outlook](../outlook/insert-data-in-the-body.md)
    
- [Obtener o definir la ubicación al redactar una cita en Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Obtener o establecer la hora al redactar una cita en Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
