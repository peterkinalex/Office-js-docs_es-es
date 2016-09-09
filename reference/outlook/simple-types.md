

# Tipos simples

####  AsyncResult

Un objeto que encapsula el resultado de una solicitud asincrónica, incluida la información de estado y error si la solicitud falla.

##### Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`asyncContext`| Object|Obtiene el elemento que se pasa al parámetro `asyncContext` opcional del método invocado, en el mismo estado en el que se pasó.|
|`error`| Error|Obtiene un objeto Error que ofrece una descripción del error, en caso de existir.|
|`status`| [Office.AsyncResultStatus](Office.md#.AsyncResultStatus-string)|Obtiene el estado de la acción asíncrona.|
|`value`| Object|Obtiene la carga o el contenido de una operación asincrónica, si la hay.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|Modo de Outlook aplicable| Redacción o lectura|
#### AttachmentDetails

Representa datos adjuntos en un elemento del servidor. Solo modo Lectura.

Una matriz de objetos `AttachmentDetail` se devuelve como la propiedad `attachments` de un objeto `Appointment` o `Message`.

##### Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`attachmentType`| [Office.MailboxEnums.AttachmentType](Office.MailboxEnums.md#attachmenttype-string)|Obtiene un valor que indica el tipo de datos adjuntos.|
|`contentType`| String|Obtiene el tipo de contenido MIME de los datos adjuntos.|
|`id`| String|Obtiene el identificador de datos de adjuntos de Exchange de los datos adjuntos.|
|`isInline`| Boolean|Obtiene un valor que indica si se deben mostrar los datos adjuntos en el cuerpo del elemento.|
|`name`| String|Obtiene el nombre de los datos adjuntos|
|`size`| Número|Obtiene el tamaño de los datos adjuntos en bytes.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|
#### Contacto

Representa un contacto almacenado en el servidor. Solo modo Lectura.

La lista de contactos asociados a un mensaje de correo electrónico o a una cita se devuelve en la propiedad `contacts` del objeto [`Entities`](simple-types.md#entities) que se devuelve mediante el método `getEntities` o `getEntitiesByType` del elemento activo.

##### Propiedades:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;nullable&gt;|Una matriz de cadenas que contiene las direcciones de correo electrónico y postales asociadas al contacto.|
|`businessName`| String| &lt;nullable&gt;|Cadena que contiene el nombre de la empresa asociada al contacto.|
|`emailAddresses`| Array.&lt;String&gt;| &lt;nullable&gt;|Una matriz de cadenas que contiene las direcciones de correo electrónico SMTP asociadas al contacto.|
|`personName`| String| &lt;nullable&gt;|Cadena que contiene el nombre de la persona asociada al contacto.|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;nullable&gt;|Una matriz que contiene un objeto `PhoneNumber` para cada número de teléfono asociado al contacto.|
|`urls`| Array.&lt;String&gt;| &lt;nullable&gt;|Una matriz de cadenas que contiene las direcciones URL de Internet asociadas al contacto.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Lectura|
####  EmailAddressDetails

Proporciona las propiedades de correo electrónico del remitente o los destinatarios especificados de un mensaje de correo o una cita.

##### Tipo:

*   Objeto

##### Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`appointmentResponse`| [Office.MailboxEnums.ResponseType](Office.MailboxEnums.md#responsetype-string)|Obtiene la respuesta que ha devuelto un asistente de una cita. Esta propiedad se aplica solo a un asistente de una cita, como se ha representado en la propiedad [`optionalAttendees`](Office.context.mailbox.item.md#optionalattendees-arrayemailaddressdetails) o [`requiredAttendees`](Office.context.mailbox.item.md#requiredattendees-arrayemailaddressdetailsrecipients). Esta propiedad devuelve `undefined` en otros escenarios.|
|`displayName`| String|Obtiene el nombre para mostrar asociado a una dirección de correo electrónico.|
|`emailAddress`| String|Obtiene la dirección de correo electrónico SMTP.|
|`recipientType`| [Office.MailboxEnums.RecipientType](Office.MailboxEnums.md#recipienttype-string)|Obtiene el tipo de dirección de correo electrónico de un destinatario.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|
#### EmailUser

Representa una cuenta de correo electrónico en Exchange Server.

##### Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`displayName`| String|Obtiene el nombre para mostrar asociado a una dirección de correo electrónico.|
|`emailAddress`| String|Obtiene la dirección de correo electrónico SMTP.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

#### Entidades

Representa una colección de entidades presente en un mensaje de correo electrónico o en una cita. Solo modo Lectura.

El objeto `Entities` es un contenedor para las matrices de entidades devueltas por los métodos `getEntities` y `getEntitiesByType` cuando el elemento (un mensaje de correo electrónico o una cita) contiene una o varias entidades encontradas por el servidor. Puede usar estas entidades en el código para proporcionar información de contexto adicional al visor, como un mapa para una dirección encontrada en el elemento o para abrir un marcador para un número de teléfono encontrado en el elemento.

Si no existen entidades del tipo especificado en la propiedad en el elemento, la propiedad asociada a esa entidad es `null`. Por ejemplo, si un mensaje contiene una dirección postal y un número de teléfono, las propiedades `addresses` y `phoneNumbers` tendrán información y las demás propiedades serán `null`.

Para que se reconozca como una dirección, la cadena debe contener una dirección postal de los Estados Unidos que tenga al menos un subconjunto de elementos con el número de la calle, el nombre de la calle, la ciudad, el estado y el código postal.

Para que se reconozca como un número de teléfono, la cadena debe contener un formato de número de teléfono de Estados Unidos.

El reconocimiento de entidades depende del reconocimiento del lenguaje natural basado en el aprendizaje automático de grandes volúmenes de datos. El reconocimiento de una entidad no es determinista y a veces los resultados dependen del contexto concreto del elemento.

Cuando es el método `getEntitiesByType` el que devuelve las matrices de propiedades, solo la propiedad de la entidad especificada contiene datos; todas las demás propiedades son `null`.

##### Propiedades:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;nullable&gt;|Obtiene las direcciones físicas (dirección postal o de correo) presentes en un mensaje de correo o una cita.|
|`contacts`| Array.&lt;[Contact](simple-types.md#contact)&gt;| &lt;nullable&gt;|Obtiene los contactos presentes en una dirección de correo electrónico o una cita.|
|`emailAddresses`| Array.&lt;String&gt;| &lt;nullable&gt;|Obtiene las direcciones de correo electrónico presentes en un mensaje de correo o una cita.|
|`meetingSuggestions`| Array.&lt;[MeetingSuggestion](simple-types.md#meetingsuggestion)&gt;| &lt;nullable&gt;|Obtiene las sugerencias de reunión presentes en un mensaje de correo.|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;nullable&gt;|Obtiene los números de teléfono presentes en un mensaje de correo o una cita.|
|`taskSuggestions`| Array.&lt;[TaskSuggestion](simple-types.md#tasksuggestion)&gt;| &lt;nullable&gt;|Obtiene las sugerencias de tarea presentes en un mensaje de correo o una cita.|
|`urls`| Array.&lt;String&gt;| &lt;nullable&gt;|Obtiene las direcciones URL de Internet presentes en un mensaje de correo o una cita.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|
#### LocalClientTime

Representa una fecha y una hora en la zona horaria local del cliente. Solo modo Lectura.

##### Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`month`| Número|Valor entero que representa el mes, comenzando con 0 para enero hasta 11 para diciembre.|
|`date`| Número|Valor entero que representa el día del mes.|
|`year`| Número|Valor entero que representa el año.|
|`hours`| Número|Valor entero que representa la hora en un reloj de 24 horas.|
|`minutes`| Número|Valor entero que representa los minutos.|
|`seconds`| Número|Valor entero que representa los segundos.|
|`milliseconds`| Número|Valor entero que representa los milisegundos.|
|`timezoneOffset`| Número|Valor entero que representa el número de minutos de diferencia entre la zona horaria local y la hora UTC.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|
#### MeetingSuggestion

Representa una reunión sugerida encontrada en un elemento. Solo modo Lectura.

La lista de reuniones sugeridas en un mensaje de correo electrónico se devuelve en la propiedad `meetingSuggestions` del objeto [`Entities`](simple-types.md#entities) que se devuelve cuando se llama al método [`getEntities`](Office.context.mailbox.item.md#getentities--entities) o [`getEntitiesByType`](Office.context.mailbox.item.md#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) en el elemento activo.

Los valores `start` y `end` son representaciones de cadenas de un objeto Date que contiene la fecha y la hora en la que la reunión sugerida va a comenzar y finalizar. Los valores se encuentran en la zona horaria predeterminada especificada para el usuario actual.

##### Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`attendees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|Obtiene los asistentes de una reunión sugerida.|
|`end`| String|Obtiene la fecha y la hora en que finalizará una reunión sugerida.|
|`location`| String|Obtiene la ubicación de una reunión sugerida.|
|`meetingString`| String|Obtiene una cadena identificada como sugerencia de reunión.|
|`start`| String|Obtiene la fecha y hora en que empezará una reunión sugerida.|
|`subject`| String|Obtiene el asunto de una reunión sugerida.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|
####  NotificationMessageDetails

Una matriz de objetos `NotificationMessageDetails` que se devuelven mediante el método [`NotificationMessages.getAllAsync`](NotificationMessages.md#getallasyncoptions-callback).

##### Tipo:

*   Objeto

##### Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`key`| String|El identificador para el mensaje de notificación.|
|`type`| [Office.MailboxEnums.ItemNotificationMessageType](Office.MailboxEnums.md#.ItemNotificationMessageType)|El tipo de mensaje de notificación.|
|`icon`| String|El identificador de recurso del icono que se usa para el mensaje. Solo se aplica cuando `type` es `InformationalMessage`.|
|`message`| String|Este es el texto del mensaje. La longitud máxima es de 150 caracteres.|
|`persistent`| Boolean|Si `true`, el mensaje permanece hasta que el complemento lo quita o el usuario lo descarta. Si `false`, se quita cuando el usuario se desplaza a un elemento diferente. Solo se aplica cuando `type` es `InformationalMessage`.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|
#### PhoneNumber

Representa un número de teléfono identificado en un elemento. Solo modo Lectura.

Una matriz de objetos `PhoneNumber` que contiene los números de teléfono presentes en un mensaje de correo electrónico se devuelve en la propiedad `phoneNumbers` del objeto [`Entities`](simple-types.md#entities) que se devuelve al llamar al método [`getEntities`](Office.context.mailbox.item.md#getEntities) en el elemento seleccionado.

##### Tipo:

*   Objeto

##### Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`originalPhoneString`| String|Obtiene el texto que se identificó en un elemento como número de teléfono.|
|`phoneString`| String|Obtiene una cadena que contiene un número de teléfono. Esta cadena contiene solo los dígitos del número de teléfono y excluye caracteres como paréntesis y guiones si estos aparecen en el elemento original.|
|`type`| String|Obtiene una cadena que identifica el tipo de número de teléfono: `Home`, `Work`, `Mobile`, `Unspecified`.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|
#### TaskSuggestion

Representa una tarea sugerida identificada en un elemento. Solo modo Lectura.

La lista de tareas sugeridas en un mensaje de correo electrónico se devuelve en la propiedad `taskSuggestions` del objeto [`Entities`][`Entities`](simple-types.md#entities) que se devuelve cuando se llama al método [`getEntities`](Office.context.mailbox.item.md#getEntities) o [`getEntitiesByType`](Office.context.mailbox.item.md#getEntitiesByType) en el elemento activo.

##### Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`assignees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|Obtiene los usuarios a los que se debe asignar una tarea sugerida.|
|`taskString`| String|Obtiene el texto de un elemento identificado como sugerencia de tarea.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|
