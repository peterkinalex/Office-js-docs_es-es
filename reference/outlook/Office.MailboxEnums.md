 

# <a name="mailboxenums"></a>MailboxEnums

## [Office](Office.md). MailboxEnums

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción o lectura|

### <a name="members"></a>Miembros

#### <a name="attachmenttype-:string"></a>AttachmentType :String

Especifica el tipo de datos adjuntos.

AttachmentType

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`File`| String|`file`|Los datos adjuntos son un archivo.|
|`Item`| String|`item`|Los datos adjuntos son un elemento de Exchange.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción o lectura|
#### <a name="entitytype-:string"></a>EntityType :String

Especifica un tipo de entidad.

EntityType

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`Address`| String|`address`|Especifica que la entidad es una dirección postal.|
|`Contact`| String|`contact`|Especifica que la entidad es un contacto.|
|`EmailAddress`| String|`emailAddress`|Especifica que la entidad es una dirección de correo electrónico SMTP.|
|`MeetingSuggestion`| String|`meetingSuggestion`|Especifica que la entidad es una sugerencia de reunión.|
|`PhoneNumber`| String|`phoneNumber`|Especifica que la entidad es un número de teléfono de EE. UU.|
|`TaskSuggestion`| String|`taskSuggestion`|Especifica que la entidad es una sugerencia de tarea.|
|`URL`| String|`url`|Especifica que la entidad es una dirección URL de Internet.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción o lectura|
#### <a name="itemnotificationmessagetype-:string"></a>ItemNotificationMessageType :String

Especifica el tipo de mensaje de notificación de una cita o un mensaje.

ItemNotificationMessageType

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`ProgressIndicator`| String|`progressIndicator`|notificationMessage es un indicador de progreso.|
|`InformationalMessage`| String|`informationalMessage`|notificationMessage es un mensaje informativo.|
|`ErrorMessage`| String|`errorMessage`|notificationMessage es un mensaje de error.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|Modo de Outlook aplicable| Redacción o lectura|
#### <a name="itemtype-:string"></a>ItemType :String

Especifica el tipo de un elemento.

ItemType

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`Message`| String|`message`|Un mensaje de correo electrónico o una convocatoria, respuesta o cancelación de una reunión.|
|`Appointment`| String|`appointment`|Un elemento de cita.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción o lectura|
#### <a name="recipienttype-:string"></a>RecipientType :String

Especifica el tipo de destinatario de una cita.

RecipientType

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`Other`| String|`other`|El destinatario no es uno de los otros tipos de destinatario.|
|`DistributionList`| String|`distributionList`|El destinatario es una lista de distribución que contiene una lista de direcciones de correo electrónico.|
|`User`| String|`user`|El destinatario es una dirección de correo electrónico SMTP que se encuentra en el servidor Exchange.|
|`ExternalUser`| String|`externalUser`|El destinatario es una dirección de correo electrónico SMTP que no se encuentra en el servidor Exchange.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|Modo de Outlook aplicable| Redacción o lectura|
#### <a name="responsetype-:string"></a>ResponseType :String

Especifica el tipo de respuesta para la invitación a una reunión.

ResponseType

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`None`| String|`none`|No ha habido respuesta del asistente.|
|`Organizer`| String|`organizer`|El asistente es el organizador de la reunión.|
|`Tentative`| String|`tentative`|El asistente ha aceptado provisionalmente la convocatoria de reunión.|
|`Accepted`| String|`accepted`|El asistente ha aceptado la convocatoria de reunión.|
|`Declined`| String|`declined`|El asistente ha rechazado la convocatoria de reunión.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción o lectura|

#### <a name="restversion-:string"></a>RestVersion :String

Especifica la versión de la API de REST que corresponde a un identificador de elemento con formato REST. 

RestVersion

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`v1_0`| String|`v1.0`|Versión 1.0.|
|`v2_0`| String|`v2.0`|Versión 2.0.|
|`Beta`| String|`beta`|Beta.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|Modo de Outlook aplicable| Redacción o lectura|
