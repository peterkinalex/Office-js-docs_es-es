

# <a name="item"></a>elemento

## <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).item

El espacio de nombres `item` se usa para acceder al mensaje seleccionado actualmente, a una convocatoria de reunión o a una cita. Puede determinar el tipo del `item` mediante la propiedad [itemType](Office.context.mailbox.item.md#itemtype-officemailboxenumsitemtype).

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|

### <a name="example"></a>Ejemplo

En el siguiente ejemplo de código de JavaScript, se muestra cómo tener acceso a la propiedad `subject` del elemento actual en Outlook.

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a>Miembros

#### <a name="attachments-arrayattachmentdetailssimple-typesmdattachmentdetails"></a>attachments :Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

Obtiene una matriz de datos adjuntos para el elemento. Solo modo Lectura.

##### <a name="type"></a>Tipo:

*   Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="example"></a>Ejemplo

El siguiente código crea una cadena HTML con los detalles de todos los datos adjuntos en el elemento actual.

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsrecipientsmd"></a>bcc :[Recipients](Recipients.md)

Obtiene o establece los destinatarios en la línea CCO (copia carbón oculta) de un mensaje. Solo modo  Redacción.

##### <a name="type"></a>Tipo:

*   [Destinatarios](Recipients.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción|

##### <a name="example"></a>Ejemplo

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodybodymd"></a>body :[Body](Body.md)

Obtiene un objeto que proporciona métodos para manipular el cuerpo de un elemento.

##### <a name="type"></a>Tipo:

*   [Cuerpo](Body.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|
####  cc :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Destinatarios](Recipients.md)

Obtiene o establece los destinatarios en copia de un mensaje.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `cc` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **CC** del mensaje. La colección está limitada a un máximo de 100 miembros.

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `cc` devuelve un objeto `Recipients` que proporciona métodos para manipular los destinatarios en la línea **CC** del mensaje.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>(nullable) conversationId :String

Obtiene un identificador de la conversación de correo electrónico que contiene un mensaje determinado.

Puede obtener un entero para esta propiedad si la aplicación de correo se activa en respuestas o formularios de lectura o en formularios de redacción. Si, después, el usuario cambia el asunto del mensaje de respuesta, al enviar dicha respuesta, el Id. de conversación de ese mensaje cambiará y el valor que se obtuvo anteriormente ya no será de aplicación.

Obtiene un valor NULL para esta propiedad para un nuevo elemento de un formulario de redacción. Si el usuario establece un asunto y guarda el elemento, la propiedad `conversationId` devolverá un valor.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|
#### <a name="datetimecreated-date"></a>dateTimeCreated :Date

Obtiene la fecha y hora de creación de un elemento. Solo modo Lectura.

##### <a name="type"></a>Tipo:

*   Fecha

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="example"></a>Ejemplo

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified :Date

Obtiene la fecha y hora en que se modificó por última vez un elemento. Solo modo Lectura.

> **Nota:** Este miembro no se admite en Outlook para iOS ni en Outlook para Android.

##### <a name="type"></a>Tipo:

*   Fecha

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="example"></a>Ejemplo

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimetimemd"></a>end :Date|[Time](Time.md)

Obtiene o establece la fecha y la hora de finalización de la cita.

La propiedad `end` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para convertir el valor de la propiedad final a la fecha y hora local del cliente.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `end` devuelve un objeto `Date`.

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `end` devuelve un objeto `Time`.

Si usa el método [`Time.setAsync`](Time.md#setasyncdatetime-options-callback) para establecer la hora de finalización, use el método [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.

##### <a name="type"></a>Tipo:

*   Date | [Time](Time.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

En el ejemplo siguiente, se establece la hora de finalización de una cita en el modo Redacción mediante el método [`setAsync`](Time.md#setasyncdatetime-options-callback) del objeto `Time`.

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailssimple-typesmdemailaddressdetails"></a>from :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Obtiene la dirección de correo electrónico del remitente de un mensaje. Solo modo Lectura.

Las propiedades `from` y [`sender`](Office.context.mailbox.item.md#sender-emailaddressdetails) representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.

> **Nota**: La propiedad `recipientType` del objeto `EmailAddressDetails` de la propiedad `from` está `undefined`.

##### <a name="type"></a>Tipo:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|
#### <a name="internetmessageid-string"></a>internetMessageId :String

Obtiene el identificador de mensaje de Internet de un mensaje de correo electrónico. Solo modo Lectura.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="example"></a>Ejemplo

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass :String

Obtiene la clase de elemento Servicios Web Exchange del elemento seleccionado. Solo modo Lectura.

La propiedad `itemClass` especifica la clase de mensaje del elemento seleccionado. Las siguientes son las clases de mensaje predeterminadas para el elemento de mensaje o cita.

| Tipo | Descripción | Clase de elemento |
| --- | --- | --- |
| Elementos de cita | Estos son los elementos de calendario de la clase de elemento `IPM.Appointment` o `IPM.Appointment.Occurence`. | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| Elementos de mensaje | Entre estos se encuentran los mensajes de correo electrónico que tienen la clase de mensaje predeterminada `IPM.Note` y las convocatorias de reunión, respuestas y cancelaciones que usan `IPM.Schedule.Meeting` como clase de mensaje base. | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

Puede crear clases de mensaje personalizadas para ampliar una clase de mensaje predeterminada. Por ejemplo, puede crear una clase de mensaje de cita personalizada `IPM.Appointment.Contoso`.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="example"></a>Ejemplo

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId :String

Obtiene el identificador de elemento de Servicios Web Exchange para el elemento actual. Solo modo Lectura.

El identificador que devuelve la propiedad `itemId` es el mismo que el identificador de elemento de Servicios Web Exchange. La propiedad `itemId` no es idéntica al identificador de entrada de Outlook.

La propiedad `itemId` no está disponible en el modo redacción. Si se requiere un identificador de elemento, el método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) puede usarse para guardar el elemento en el servidor, que devolverá el identificador de elemento en el parámetro [`AsyncResult.value`](simple-types.md#asyncresult) de la función de devolución de llamada.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="example"></a>Ejemplo

El siguiente código comprueba la presencia de un identificador de elemento. Si la propiedad `itemId` devuelve `null` o `undefined`, guarda el elemento en el servidor y obtiene el identificador de elemento desde el resultado asincrónico.

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypeofficemailboxenumsmditemtype-string"></a>itemType :[Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

Obtiene el tipo de elemento que representa una instancia.

La propiedad `itemType` devuelve uno de los valores de enumeración de `ItemType`, lo que indica si la instancia del objeto `item` es un mensaje o una cita.

##### <a name="type"></a>Tipo:

*   [Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationlocationmd"></a>location :String|[Location](Location.md)

Obtiene o establece la ubicación de una cita.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `location` devuelve una cadena que contiene la ubicación de la cita.

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `location` devuelve un objeto `Location` que proporciona métodos para obtener y establecer la ubicación de la cita.

##### <a name="type"></a>Tipo:

*   String | [Location](Location.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject :String

Obtiene el asunto de un elemento, con todos los prefijos quitados (incluidos `RE:` y `FWD:`). Solo modo Lectura.

La propiedad normalizedSubject obtiene el asunto del elemento, con los prefijos estándar (como `RE:` y `FW:`) agregados por los programas de correo electrónico. Para obtener el asunto del elemento con los prefijos intactos, use la propiedad [`subject`](Office.context.mailbox.item.md#subject-stringsubject).

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="example"></a>Ejemplo

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesnotificationmessagesmd"></a>notificationMessages :[NotificationMessages](NotificationMessages.md)

Obtiene los mensajes de notificación de un elemento.

##### <a name="type"></a>Tipo:

*   [NotificationMessages](NotificationMessages.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|
|[Destinatarios](Recipients.md)|
####  <a name="optionalattendees-arrayemailaddressdetailssimple-typesmdemailaddressdetails"></a>optionalAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>
Obtiene o establece una lista de direcciones de correo para los asistentes opcionales.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `optionalAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente opcional a la reunión.

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `optionalAttendees` devuelve un objeto `Recipients` que proporciona métodos para obtener y establecer los asistentes opcionales para una reunión.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailssimple-typesmdemailaddressdetails"></a>organizer :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Obtiene la dirección de correo electrónico del organizador de una reunión especificada. Solo modo Lectura.

##### <a name="type"></a>Tipo:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="example"></a>Ejemplo

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailssimple-typesmdemailaddressdetailsrecipientsrecipientsmd"></a>requiredAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

Obtiene o define una lista de direcciones de correo electrónico para los asistentes necesarios.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `requiredAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente requerido a la reunión.

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `requiredAttendees` devuelve un objeto `Recipients` que proporciona métodos para obtener y establecer los asistentes requeridos para una reunión.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="resources-emailaddressdetailssimple-typesmdemailaddressdetails"></a>resources :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Obtiene los recursos requeridos para una cita. Solo modo Lectura.

> **Nota:** Este miembro no se admite en Outlook para iOS ni en Outlook para Android.

##### <a name="type"></a>Tipo:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|
#### <a name="sender-emailaddressdetailssimple-typesmdemailaddressdetails"></a>sender :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Obtiene la dirección de correo electrónico del remitente de un mensaje de correo electrónico. Solo modo Lectura.

Las propiedades [`from`](Office.context.mailbox.item.md#from-emailaddressdetails) y `sender` representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.

> **Nota**: La propiedad `recipientType` del objeto `EmailAddressDetails` de la propiedad `sender` está `undefined`.

##### <a name="type"></a>Tipo:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="example"></a>Ejemplo

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimetimemd"></a>start :Date|[Time](Time.md)

Obtiene o establece la fecha y la hora de inicio de la cita.

La propiedad `start` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para convertir el valor a la fecha y hora local del cliente.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `start` devuelve un objeto `Date`.

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `start` devuelve un objeto `Time`.

Si usa el método [`Time.setAsync`](Time.md#setasyncdatetime-options-callback) para establecer la hora de inicio, use el método [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.

##### <a name="type"></a>Tipo:

*   Date | [Time](Time.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

En el ejemplo siguiente, se establece la hora de inicio de una cita en el modo Redacción mediante el método [`setAsync`](Time.md#setasyncdatetime-options-callback) del objeto `Time`.

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectsubjectmd"></a>subject :String|[Subject](Subject.md)

Obtiene o establece la descripción que se muestra en el campo de asunto de un elemento.

La propiedad `subject` obtiene o establece el asunto completo del elemento, como lo ha enviado el servidor de correo electrónico.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `subject` devuelve una cadena. Use la propiedad [`normalizedSubject`](Office.context.mailbox.item.md#normalizedsubject-string) para obtener el asunto menos cualquier prefijo inicial como `RE:` y `FW:`.

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `subject` devuelve un objeto `Subject` que proporciona métodos para obtener y establecer el asunto.

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>Tipo:

*   String | [Subject](Subject.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|
####  para :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Destinatarios](Recipients.md)

Obtiene o establece los destinatarios de un mensaje de correo.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `to` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **Para** del mensaje. La colección está limitada a un máximo de 100 miembros.

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `to` devuelve un objeto `Recipients` que proporciona métodos para manipular los destinatarios en la línea **Para** del mensaje.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>Métodos

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Agrega un archivo a un mensaje o cita como datos adjuntos.

El método `addFileAttachmentAsync` carga el archivo en el URI especificado y lo asocia al elemento en el formulario de redacción.

Después, puede usar el identificador con el método [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.

##### <a name="parametersremoveattachmentasyncattachmentid-options-callback"></a>Parameters:removeattachmentasyncattachmentid-options-callback
|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`uri`| String||El URI que proporciona la ubicación del archivo que se va a adjuntar al mensaje o a la cita. La longitud máxima es de 2048 caracteres.|
|`attachmentName`| String||El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.<br/>Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td><code>AttachmentSizeExceeded</code></td><td>Los datos adjuntos son más grandes de lo permitido.</td></tr><tr><td><code>FileTypeNotSupported</code></td><td>Los datos adjuntos tienen una extensión que no está permitida.</td></tr><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>El mensaje o cita tiene demasiados datos adjuntos.</td></tr></tbody></table>|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|

##### <a name="example"></a>Ejemplo

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Agrega un elemento de Exchange (por ejemplo, un mensaje) como datos adjuntos al mensaje o a la cita.

El método `addItemAttachmentAsync` asocia el elemento con el identificador especificado de Exchange al elemento en el formulario de redacción. Si se especifica un método de devolución de llamada, se llama al método con un parámetro, `asyncResult`, que contiene el identificador de datos adjuntos o un código que indique cualquier error que se ha producido al adjuntar el elemento. Si es necesario, puede usar el parámetro `options` para pasar información de estado al método de devolución de llamada.

Después, puede usar el identificador con el método [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.

Si el complemento de Office se ejecuta en Outlook Web App, el método `addItemAttachmentAsync` puede adjuntar elementos a otros elementos distintos del elemento que quiere editar. No obstante, esta acción no es compatible y no se recomienda.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`itemId`| String||El identificador de Exchange del elemento que debe adjuntarse. La longitud máxima es de 100 caracteres.|
|`attachmentName`| String||El asunto del elemento que debe adjuntarse. La longitud máxima es de 255 caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.<br/>Si se produce un error al agregar los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>El mensaje o cita tiene demasiados datos adjuntos.</td></tr></tbody></table>|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|

##### <a name="example"></a>Ejemplo

En el ejemplo siguiente, se agrega un elemento existente de Outlook como un archivo adjunto con el nombre `My Attachment`.

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a>close()

Cierra el elemento actual que se está redactando.

El comportamiento del método `close` depende del estado actual del elemento que se está redactando. Si el elemento tiene cambios sin guardar, el cliente solicita al usuario que guarde, descarte o cancele la acción Cerrar.

En el cliente de escritorio de Outlook, si el mensaje es una respuesta directa, el método `close` no tiene ningún efecto.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción|
#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

Muestra un formulario de respuesta que incluye el remitente y todos los destinatarios del mensaje seleccionado o el organizador y todos los asistentes de la cita seleccionada.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.

Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyAllForm` produce una excepción.

Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`formData`| String &#124; Object|Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.<br/>**O**<br/>Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define como sigue:<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;optional&gt;</td><td>Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</td></tr><tr><td><code>attachments</code></td><td>Array.&lt;Object&gt;</td><td>&lt;optional&gt;</td><td>Una matriz de objetos JSON que son elementos o archivos adjuntos.<br/><br/><strong>Propiedades</strong><br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Descripción</th></tr></thead><tbody><tr><td><code>type</code></td><td>String</td><td>Indica el tipo de datos adjuntos. Debe ser <code>file</code> para un archivo adjunto o <code>item</code> para un elemento adjunto.</td></tr><tr><td><code>name</code></td><td>String</td><td>Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</td></tr><tr><td><code>url</code></td><td>String</td><td>Solo se usa si <code>type</code> se establece en <code>file</code>. El URI de la ubicación del archivo.</td></tr><tr><td><code>itemId</code></td><td>String</td><td>Solo se usa si <code>type</code> se establece en <code>item</code>. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</td></tr></tbody></table></td></tr><tr><td><code>callback</code></td><td>función</td><td>&lt;optional&gt;</td><td>Cuando el método finaliza, la función que se pasa en el parámetro <code>callback</code> se llama con un único parámetro, <code>asyncResult</code>, que es un objeto <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a>. Para obtener más información, consulte <a href="tutorial-asynchronous.html">Usar los métodos asincrónicos</a>.</td></tr></tbody></table>|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="examples"></a>Ejemplos

El código siguiente pasa una cadena a la función `displayReplyAllForm`.

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Responder con un cuerpo vacío.

```
Office.context.mailbox.item.displayReplyAllForm({});
```

Responder solo con un cuerpo.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Responder con un cuerpo y datos adjuntos de archivo.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Responder con un cuerpo y datos adjuntos de elemento.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a>displayReplyForm(formData)

Muestra un formulario de respuesta que incluye solo el remitente del mensaje seleccionado o el organizador de la cita seleccionada.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.

Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyForm` produce una excepción.

Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`formData`| String &#124; Object|Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.<br/>**O**<br/>Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define como sigue:<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;optional&gt;</td><td>Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</td></tr><tr><td><code>attachments</code></td><td>Array.&lt;Object&gt;</td><td>&lt;optional&gt;</td><td>Una matriz de objetos JSON que son elementos o archivos adjuntos.<br/><br/><strong>Propiedades</strong><br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Descripción</th></tr></thead><tbody><tr><td><code>type</code></td><td>String</td><td>Indica el tipo de datos adjuntos. Debe ser <code>file</code> para un archivo adjunto o <code>item</code> para un elemento adjunto.</td></tr><tr><td><code>name</code></td><td>String</td><td>Solo se usa si <code>type</code> se establece en <code>file</code>. Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</td></tr><tr><td><code>url</code></td><td>String</td><td>Solo se usa si <code>type</code> se establece en <code>file</code>. El URI de la ubicación del archivo.</td></tr><tr><td><code>itemId</code></td><td>String</td><td>Solo se usa si <code>type</code> se establece en <code>item</code>. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</td></tr></tbody></table></td></tr><tr><td><code>callback</code></td><td>función</td><td>&lt;optional&gt;</td><td>Cuando el método finaliza, la función que se pasa en el parámetro <code>callback</code> se llama con un único parámetro, <code>asyncResult</code>, que es un objeto <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a>. Para obtener más información, consulte <a href="tutorial-asynchronous.html">Usar los métodos asincrónicos</a>.</td></tr></tbody></table>|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="examples"></a>Ejemplos

El código siguiente pasa una cadena a la función `displayReplyForm`.

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Responder con un cuerpo vacío.

```
Office.context.mailbox.item.displayReplyForm({});
```

Responder solo con un cuerpo.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Responder con un cuerpo y datos adjuntos de archivo.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Responder con un cuerpo y datos adjuntos de elemento.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiessimple-typesmdentities"></a>getEntities() → {[Entities](simple-types.md#entities)}

Obtiene las entidades que se han encontrado en el elemento seleccionado.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="returns"></a>Valores devueltos:

Tipo: [Entidades](simple-types.md#entities)

##### <a name="example"></a>Ejemplo

En el ejemplo siguiente, se tiene acceso a las entidades de contactos del artículo actual.

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactsimple-typesmdcontactmeetingsuggestionsimple-typesmdmeetingsuggestionphonenumbersimple-typesmdphonenumbertasksuggestionsimple-typesmdtasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

Obtiene una matriz de todas las entidades del tipo de entidad especificado que se han encontrado en el elemento seleccionado.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](Office.MailboxEnums.md#.EntityType-string)|Uno de los valores de enumeración de EntityType.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Lectura|

##### <a name="returns"></a>Valores devueltos:

Si el valor que se pasa a `entityType` no es un miembro válido de la enumeración de `EntityType`, el método devuelve un valor NULL. Si no está presente ninguna entidad del tipo especificado en el elemento, el método devuelve una matriz vacía. De otro modo, el tipo de los objetos en la matriz devuelta depende del tipo de entidad solicitada en el parámetro `entityType`.

Aunque el nivel de permiso mínimo para usar este método es **Restricted**, algunos tipos de entidad requieren **ReadItem** para tener acceso, como se especifica en la tabla siguiente.

| Valor de `entityType` | Tipo de objetos de la matriz devuelta | Nivel de permiso requerido |
| --- | --- | --- |
| `Address` | Cadena | **Restringido** |
| `Contact` | Contacto | **ReadItem** |
| `EmailAddress` | Cadena | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **Restringido** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | Cadena | **Restringido** |

Tipo: Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>

##### <a name="example"></a>Ejemplo

En el ejemplo siguiente, se muestra cómo tener acceso a una matriz de cadenas que representan las direcciones postales en el asunto o en el cuerpo del elemento actual.

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactsimple-typesmdcontactmeetingsuggestionsimple-typesmdmeetingsuggestionphonenumbersimple-typesmdphonenumbertasksuggestionsimple-typesmdtasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

Devuelve entidades conocidas en el elemento seleccionado que pasan el filtro con nombre definido en el archivo XML de manifiesto.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

El método `getFilteredEntitiesByName` devuelve las entidades que coinciden con la expresión regular definida en el elemento de regla [ItemHasKnownEntity](https://msdn.microsoft.com/en-us/library/office/fp161166.aspx) del archivo XML de manifiesto con el valor de elemento `FilterName` especificado.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`name`| String|El nombre del elemento de regla `ItemHasKnownEntity` que define el filtro para que coincidan.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="returns"></a>Valores devueltos:

Si no existe ningún elemento `ItemHasKnownEntity` en el manifiesto con un valor de elemento `FilterName` que coincida con el parámetro `name`, el método devuelve `null`. Si el parámetro `name` no coincide con un elemento `ItemHasKnownEntity` del manifiesto, pero no existen entidades en el elemento actual que coincidan, el método devuelve una matriz vacía.

Tipo: Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>

#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

Devuelve valores de cadena en el elemento seleccionado que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

El método `getRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.

Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados. En su lugar, use el método [`Body.getAsync`](Body.md#getasynccoerciontype-options-callback) para recuperar todo el cuerpo.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="returns"></a>Valores devueltos:

Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.

<dl class="param-type">

<dt>Tipo</dt>

<dd>Object</dd>

</dl>

##### <a name="example"></a>Ejemplo

En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los <rule>elementos de expresión regular `fruits`y `veggies`, que se especifican en el manifiesto.</rule>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-arraystring"></a>getRegExMatchesByName(name) → (nullable) {Array.<String>}

Devuelve valores de cadena en el elemento seleccionado que coinciden con la expresión regular de nombre definida en el archivo XML de manifiesto.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

El método `getRegExMatchesByName` devuelve las cadenas que coinciden con la expresión regular definida en el elemento de regla `ItemHasRegularExpressionMatch` del archivo XML de manifiesto con el valor de elemento `RegExName` especificado.

Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`name`| String|El nombre del elemento de regla `ItemHasRegularExpressionMatch` que define el filtro para que coincidan.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="returns"></a>Valores devueltos:

Una matriz que contiene las cadenas que coinciden con la expresión regular definida en el archivo XML de manifiesto.

<dl class="param-type">

<dt>Tipo</dt>

<dd>Matriz.<String></dd>

</dl>

##### <a name="example"></a>Ejemplo

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

Devuelve asincrónicamente datos seleccionados desde el asunto o el cuerpo de un mensaje.

Si no hay ninguna selección, pero el cursor está en el cuerpo o el asunto, el método devuelve null para los datos seleccionados. Si se selecciona un campo que no sea el cuerpo o el asunto, el método devuelve el error `InvalidSelection`.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||Solicita un formato para los datos. Si es Texto, el método devuelve el texto sin formato como cadena, quitando toda etiqueta HTML que hubiera. Si es HTML, el método devuelve el texto seleccionado, ya sea texto sin formato o HTML.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

Para tener acceso a los datos seleccionados desde el método de devolución de llamada, llame a `asyncResult.value.data`. Para tener acceso a la propiedad de origen de la que proviene la selección, llame a `asyncResult.value.sourceProperty`, que será `body` o `subject`.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|

##### <a name="returns"></a>Valores devueltos:

Los datos seleccionados como cadena con formato determinado por `coercionType`.

<dl class="param-type">

<dt>Tipo</dt>

<dd>String</dd>

</dl>

##### <a name="example"></a>Ejemplo

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

Carga de forma asincrónica las propiedades personalizadas de este complemento en el elemento seleccionado.

Las propiedades personalizadas se almacenan como pares de clave/valor según la aplicación y el elemento. Este método devuelve un objeto `CustomProperties` en la devolución de llamada, que proporciona métodos para tener acceso a las propiedades personalizadas específicas del elemento y el complemento actual. Las propiedades personalizadas no están cifradas en el elemento, por lo que no debería usarse como almacenamiento seguro.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

Las propiedades personalizadas se proporcionan como un objeto [`CustomProperties`](CustomProperties.md) en la propiedad `asyncResult.value`. Este objeto puede usarse para obtener, establecer y eliminar las propiedades personalizadas del elemento, y guardar los cambios en la propiedad personalizada que se vuelve a establecer en el servidor.| |`userContext`| Object| &lt;optional&gt;|Los desarrolladores pueden proporcionar cualquier objeto que deseen para acceder a la función de devolución de llamada. A este objeto se puede acceder mediante la propiedad `asyncResult.asyncContext` en la función de devolución de llamada.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

En el siguiente ejemplo de código se muestra cómo usar el método `loadCustomPropertiesAsync` para cargar de forma asincrónica propiedades personalizadas que son específicas del elemento actual. En el ejemplo también se muestra cómo usar el método `CustomProperties.saveAsync` para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, el ejemplo de código usa el método `CustomProperties.get` para leer la propiedad personalizada `myProp`, el método `CustomProperties.set` para escribir en la propiedad personalizada `otherProp` y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

Quita los datos adjuntos de un mensaje o cita.

El método `removeAttachmentAsync` quita del elemento los datos adjuntos con el identificador especificado. Como práctica recomendada, debe usar el identificador de datos adjuntos para quitar datos adjuntos solo si la misma aplicación de correo ha agregado los datos adjuntos en la misma sesión. En Outlook Web App y OWA para los dispositivos, el identificador de datos adjuntos es válido solo en la misma sesión. Una sesión finaliza cuando el usuario cierra la aplicación, o si el usuario comienza a redactar en un formulario en línea y posteriormente se extrae el formulario en línea para continuar en una ventana independiente.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`attachmentId`| String||El identificador de los datos adjuntos que se deben quitar. La longitud máxima de la cadena es de 100 caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Si se produce un error en la eliminación de los datos adjuntos, la propiedad `asyncResult.error` contendrá un código de error con el motivo del error.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td><code>InvalidAttachmentId</code></td><td>El identificador de datos adjuntos no existe.</td></tr></tbody></table>|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|

##### <a name="example"></a>Ejemplo

El código siguiente quita un archivo adjunto con un identificador de "0".

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

Guarda un elemento de forma asincrónica.

Cuando se invoca, este método guarda el mensaje actual como un borrador y devuelve el identificador de elemento a través del método de devolución de llamada. En el modo en línea de Outlook Web App o Outlook, el elemento se guarda en el servidor. En modo en caché de Outlook, se guarda el elemento en la caché local.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

Si se ha realizado correctamente, el identificador de elemento se proporciona en la propiedad `asyncResult.value`.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|

##### <a name="examples"></a>Ejemplos

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

A continuación, se muestra un ejemplo del parámetro `result` que se ha pasado a la función de devolución de llamada. La propiedad `value` contiene el identificador de elemento del elemento.

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

Inserta asincrónicamente datos en el cuerpo o el asunto de un mensaje.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`data`| String||Datos que se van a insertar. Los datos no deben superar 1.000.000 de caracteres. Si se pasan más de 1.000.000 de caracteres, se produce una excepción `ArgumentOutOfRange`.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>Si <code>text</code>, se aplica el estilo actual en Outlook Web App y Outlook. Si el campo es un editor de HTML, se insertan solo los datos de texto, aunque los datos sean HTML.</td></tr></tbody></table><p>Si <code>html</code> y el campo admiten HTML (el asunto no), se aplica el estilo actual en Outlook Web App y el estilo predeterminado en Outlook. Si el campo es un campo de texto, se devuelve un error <code>InvalidDataFormat</code>.</p><p>Si <code>coercionType</code> no está establecido, el resultado depende del campo: si el campo es HTML, se usa HTML; si el campo es texto, se usa texto sin formato.|</p>|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). |

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.2|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|

##### <a name="example"></a>Ejemplo

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
