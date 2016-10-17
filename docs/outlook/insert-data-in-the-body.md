
# <a name="insert-data-in-the-body-when-composing-an-appointment-or-message-in-outlook"></a>Introducir datos en el cuerpo al redactar una cita o un mensaje en Outlook

Puede usar los métodos asincrónicos ([Body.getAsync](../../reference/outlook/Body.md), [Body.getTypeAsync](../../reference/outlook/Body.md), [Body.prependAsync](../../reference/outlook/Body.md), [Body.setAsync](../../reference/outlook/Body.md) y [Body.setSelectedDataAsync](../../reference/outlook/Body.md)) para obtener el tipo de cuerpo e insertar datos en el cuerpo de la cita o el elemento de mensaje que el usuario está redactando. Estos métodos asincrónicos solo están disponibles para complementos de redacción. Para usar estos métodos, asegúrese de configurar el manifiesto del complemento correctamente para que Outlook active el complemento en los formularios de redacción, como se describe en [Crear complementos de Outlook para formularios de redacción](../outlook/compose-scenario.md).

En Outlook, un usuario puede crear un mensaje de texto, HTML o formato de texto enriquecido (RTF) y puede crear una cita en formato HTML. Antes de insertar, siempre tiene que comprobar primero el formato de elemento compatible. Para hacerlo, llame a **getTypeAsync**, ya que es posible que tenga que realizar pasos adicionales. El valor que devuelve **getTypeAsync** depende del formato del elemento original, así como de la compatibilidad del sistema operativo del dispositivo y host para la edición en formato HTML (1). Después, establezca el parámetro _coercionType_ de **prependAsync** o **setSelectedDataAsync** en consecuencia (2) para insertar los datos, como se muestra en la tabla siguiente. Si no especifica un argumento, **prependAsync** y **setSelectedDataAsync** suponen que los datos que se insertarán están en formato de texto.



|**Datos para insertar**|**Formato de elemento devuelto por getTypeAsync**|**Usar este coercionType**|
|:-----|:-----|:-----|
|Texto|Texto (1)|Texto|
|HTML|Texto (1)|Texto (2)|
|Texto|HTML|Texto/HTML|
|HTML|HTML |HTML|

1.  En tabletas y smartphones, **getTypeAsync** devuelve **Office.MailboxEnums.BodyType.Text** si el sistema operativo o el host no admiten la edición de un elemento, que se creó originalmente en HTML, en formato HTML.

2.  Si los datos que se insertarán son HTML y **getTypeAsync** devuelve un tipo de texto para ese elemento, reorganice los datos como texto e insértelos con **Office.MailboxEnums.BodyType.Text** como _coercionType_. Si simplemente inserta los datos HTML con un tipo de coerción de texto, el host mostraría las etiquetas HTML como texto. Si intenta insertar los datos HTML con **Office.MailboxEnums.BodyType.Html** como _coercionType_, obtendrá un error.

Además de _coercionType_, como con la mayoría de los métodos asincrónicos en la API de JavaScript para Office, **getTypeAsync**, **prependAsync** y **setSelectedDataAsync** aceptan otros parámetros de entrada opcionales. Para más información sobre cómo especificar estos parámetros de entrada opcionales, vea [Pasar parámetros opcionales a métodos asincrónicos](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) en [Programación asincrónica en complementos de Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## <a name="to-insert-data-at-the-current-cursor-position"></a>Para insertar datos en la posición actual del cursor


Esta sección muestra un ejemplo de código que usa  **getTypeAsync** para verificar el tipo de cuerpo del elemento que se está redactando y, a continuación, usa **setSelectedDataAsync** para insertar datos en la ubicación actual del cursor.

Puede pasar un método de devolución de llamada y parámetros de entrada opcionales a  **getTypeAsync** y obtener el estado y los resultados en el parámetro de salida _asyncResult_. Si el método se ejecuta correctamente, puede obtener el tipo de cuerpo del elemento en la propiedad [AsyncResult.value](../../reference/shared/asyncresult.status.md), que es "text" o "html".

Debe pasar una cadena de datos como parámetro de entrada a  **setSelectedDataAsync**. De acuerdo con el tipo de cuerpo del elemento, puede especificar esta cadena de datos en formato de texto o HTML, según corresponda. Como se mencionó anteriormente, si lo desea puede especificar el tipo de datos que se insertarán en el parámetro  _coercionType_. Además, puede proporcionar un método de devolución de llamada y cualquiera de sus parámetros como parámetros de entrada opcionales.

Si el usuario aún no ha colocado el cursor en el cuerpo del elemento,  **setSelectedDataAsync** insertará los datos en la parte superior del cuerpo. Si el usuario ha seleccionado texto en el cuerpo del elemento, **setSelectedDataAsync** reemplazará el texto seleccionado con los datos que especifique. Tenga en cuenta que **setSelectedDataAsync** puede dar error si el usuario está cambiando simultáneamente la posición del cursor mientras redacta el elemento. El número máximo de caracteres que puede insertar de una vez es de 1.000.000.

Este código de ejemplo asume una regla en el manifiesto del complemento que activa el complemento en un formulario de redacción para una cita o un mensaje, tal como se muestra a continuación.




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set data in the body of the composed item.
        setItemBody();
    });
}


// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.setSelectedDataAsync(
                        '<b> Kindly note we now open 7 days a week.</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.setSelectedDataAsync(
                        ' Kindly note we now open 7 days a week.',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="to-insert-data-at-the-beginning-of-the-item-body"></a>Para insertar datos al comienzo del cuerpo del elemento


Asimismo, también puede usar  **prependAsync** para insertar datos al comienzo del cuerpo del elemento e ignorar la posición actual del cursor. Excepto por el punto de inserción, **prependAsync** y **setSelectedDataAsync** funcionan de manera similar:


- Si antepone datos HTML en el cuerpo de un mensaje, primero debe comprobar el tipo del cuerpo del mensaje para no anteponer datos HTML en un mensaje con formato de texto.
    
- Proporcione los siguientes valores como parámetros de entrada de  **prependAsync**: una cadena de datos en formato de texto o HTML y, opcionalmente, el formato de los datos que se van a insertar, un método de devolución de llamada y cualquiera de sus parámetros.
    
- El número máximo de caracteres que se pueden anteponer de una vez es de 1.000.000.
    
El siguiente código JavaScript es parte de un complemento de muestra que se activa en los formularios de redacción de citas y mensajes. El ejemplo llama a  **getTypeAsync** para verificar el tipo del cuerpo del elemento e inserta datos HTML en la parte superior del cuerpo del elemento si el elemento es una cita o un mensaje HTML; de lo contrario, inserta los datos en formato de texto.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        prependItemBody();
    });
}

// Get the body type of the composed item, and prepend data  
// in the appropriate data type in the item body.
function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings!</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.prependAsync(
                        'Greetings!',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
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
    
- [Obtener o establecer el asunto al redactar una cita o un mensaje en Outlook](../outlook/get-or-set-the-subject.md)
    
- [Obtener o definir la ubicación al redactar una cita en Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Obtener o establecer la hora al redactar una cita en Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
