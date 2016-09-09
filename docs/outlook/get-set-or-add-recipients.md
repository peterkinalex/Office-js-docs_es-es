
# Obtener, establecer o agregar destinatarios al redactar una cita o un mensaje en Outlook


La API de JavaScript para Office ofrece métodos asincrónicos ([Recipients.getAsync](../../reference/outlook/Recipients.md), [Recipients.setAsync](../../reference/outlook/Recipients.md) o [Recipients.addAysnc](../../reference/outlook/Recipients.md)) para, respectivamente, obtener, definir o agregar destinatarios en un formulario de redacción de una cita o un mensaje. Estos métodos asincrónicos solo están disponibles para complementos de redacción. Para usar estos métodos, asegúrese de configurar el manifiesto del complemento correctamente para que Outlook active el complemento en los formularios de redacción, como se describe en la sección [Crear complementos de Outlook para formularios de redacción](../outlook/compose-scenario.md).

Algunas de las propiedades que representan destinatarios de una cita o mensaje están disponibles para el acceso de lectura tanto en formularios de redacción como en formularios de lectura. Entre estas propiedades se incluyen [optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md) y [requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md) para citas y [cc](../../reference/outlook/Office.context.mailbox.item.md) y [to](../../reference/outlook/Office.context.mailbox.item.md) para mensajes. En los formularios de lectura puede obtener acceso a la propiedad directamente desde el objeto primario, por ejemplo:




```js
item.cc
```

Sin embargo, en los formularios de redacción tanto el usuario como el complemento pueden insertar o cambiar un destinatario al mismo tiempo. Por ello, es necesario usar el método asincrónico  **getAsync** para obtener dichas propiedades, como en el siguiente ejemplo:




```js
item.cc.getAsync
```

Estas propiedades están disponibles para acceso de escritura en formularios que son solamente de redacción, no de lectura.

Como con la mayoría de los métodos asincrónicos en la API de JavaScript para Office, **getAsync**, **setAsync** y **addAsync** aceptan parámetros de entrada opcionales. Para más información sobre la especificación de estos parámetros de entrada opcionales, vea [Pasar parámetros opcionales a métodos asincrónicos](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) en [Programación asincrónica en complementos de Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## Para obtener los destinatarios


En esta sección se muestra un ejemplo de código que obtiene los destinatarios de la cita o mensaje que se redacte y muestra sus direcciones de correo electrónico. El ejemplo de código presupone una regla en el manifiesto del complemento que activa el complemento en un formulario de redacción para una cita o mensaje, tal y como se muestra a continuación. 


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

Como en la API de JavaScript para Office las propiedades que representan a los destinatarios de una cita ( **optionalAttendees** y **requiredAttendees**) son distintas que las de un mensaje ([bcc](../../reference/outlook/Office.context.mailbox.item.md),  **cc** y **to**), primero debe usar la propiedad [item.itemType](../../reference/outlook/Office.context.mailbox.item.md) para identificar si el elemento que se está redactando es una cita o un mensaje. En el modo de redacción, todas las propiedades de las citas y los mensajes son objetos de [Recipients](../../reference/outlook/Recipients.md), por lo que se puede aplicar el método asincrónico,  **Recipients.getAsync**, para obtener los destinatarios correspondientes. 

Para usar  **getAsync**, proporcione un método de devolución de llamada para comprobar el estado, los resultados y los errores que devuelve la llamada asincrónica a  **getAsync**. Puede proporcionar los argumentos necesarios al método de devolución de llamada usando el parámetro opcional  _asyncContext_. El método de devolución de llamada devolverá un parámetro de salida  _asyncResult_. Puede usar las propiedades  **status** y **error** del objeto de parámetro [AsyncResult](../../reference/outlook/simple-types.md) para comprobar el estado y los mensajes de error de la llamada asincrónica y la propiedad **value** para obtener los destinatarios. Los destinatarios se representan como una matriz de objetos [EmailAddressDetails](../../reference/outlook/simple-types.md).

Tenga en cuenta que, puesto que el método  **getAsync** es asincrónico, si existen acciones posteriores que dependen de que la lista de destinatarios se obtenga correctamente, deberá organizar el código de forma que dichas acciones se inicien únicamente en el método de devolución de llamada correspondiente una vez que la llamada asincrónica se haya completado de manera correcta.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Para definir los destinatarios


En esta sección se muestra un ejemplo de código que define los destinatarios de la cita o mensaje que el usuario está redactando. Al definir los destinatarios, se sobrescribirán los destinatarios existentes. Al igual que en el ejemplo anterior, que obtiene los destinatarios de un formulario de redacción, este ejemplo presupone que el complemento se activa en formularios de redacción para mensajes y citas. En este ejemplo se verifica en primer lugar si el elemento redactado es una cita o un mensaje a fin de aplicar el método asincrónico  **Recipients.setAsync** en las propiedades adecuadas que representan los destinatarios de la cita o mensaje.

Al llamar a  **setAsync**, proporcione una matriz como argumento de entrada para el parámetro  _recipients_ en uno de los siguientes formatos:


- Una matriz de cadenas que son direcciones SMTP.
    
- Una matriz de diccionarios, cada uno de los cuales contiene un nombre para mostrar y una dirección de correo electrónico, tal y como se muestra en el ejemplo de código siguiente.
    
- Una matriz de objetos  **EmailAddressDetails** similar a la devuelta por el método **getAsync**.
    
Si lo desea, puede proporcionar un método de devolución de llamada como argumento de entrada al método  **setAsync** para asegurarse de que el código que dependa de que los destinatarios se definan correctamente, solamente se ejecutará cuando esto ocurra. También puede proporcionar argumentos para el método de devolución de llamada usando el parámetro opcional _asyncContext_. Si usa un método de devolución de llamada, puede tener acceso a un parámetro de salida  _asyncResult_ y usar las propiedades **status** y **error** del objeto de parámetro **AsyncResult** para comprobar el estado de los mensajes de error de la llamada asincrónica.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## Para agregar destinatarios


Si no desea sobrescribir los destinatarios existentes en una cita o mensaje, en lugar de usar  **Recipients.setAsync**, puede usar el método asincrónico  **Recipients.addAsync** para anexar destinatarios. El funcionamiento de **addAsync** se asemeja a **setAsync** en que requiere de un argumento de entrada _recipients_. Si lo desea, puede proporcionar un método de devolución de llamada y argumentos para la devolución de llamada mediante el parámetro asyncContext. A continuación, puede comprobar el estado, el resultado y los errores de la llamada asincrónica a  **addAsync** usando el parámetro de salida _asyncResult_ del método de devolución de llamada. El siguiente ejemplo comprueba si el elemento que se está redactando es una cita y le agrega dos asistentes obligatorios.


```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## Recursos adicionales



- [Obtener y definir datos de elementos en un formulario de redacción de Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Obtención y definición de datos de elementos de Outlook en los formularios de lectura o redacción](../outlook/item-data.md)
    
- [Crear complementos de Outlook para formularios de redacción](../outlook/compose-scenario.md)
    
- [Programación asíncrona en los complementos de Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Obtener o establecer el asunto al redactar una cita o un mensaje en Outlook](../outlook/get-or-set-the-subject.md)
    
- [Introducir datos en el cuerpo al redactar una cita o un mensaje en Outlook](../outlook/insert-data-in-the-body.md)
    
- [Obtener o definir la ubicación al redactar una cita en Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Obtener o establecer la hora al redactar una cita en Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
