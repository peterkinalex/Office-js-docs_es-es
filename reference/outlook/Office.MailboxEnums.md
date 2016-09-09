 

# MailboxEnums

## [Office](Office.md). MailboxEnums

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|Modo de Outlook aplicable| Redacción o lectura|

### Miembros

#### AttachmentType :String

Especifica el tipo de datos adjuntos.

AttachmentType

##### Tipo:

*   String

##### Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`File`| String|`file`|Los datos adjuntos son un archivo.|
|`Item`| String|`item`|Los datos adjuntos son un elemento de Exchange.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|Modo de Outlook aplicable| Redacción o lectura|
#### EntityType :String

Especifica un tipo de entidad.

EntityType

##### Tipo:

*   String

##### Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`Address`| String|`address`|Especifica que la entidad es una dirección postal.|
|`Contact`| String|`contact`|Especifica que la entidad es un contacto.|
|`EmailAddress`| String|`emailAddress`|Especifica que la entidad es una dirección de correo electrónico SMTP.|
|`MeetingSuggestion`| String|`meetingSuggestion`|Especifica que la entidad es una sugerencia de reunión.|
|`PhoneNumber`| String|`phoneNumber`|Especifica que la entidad es un número de teléfono de EE. UU.|
|`TaskSuggestion`| String|`taskSuggestion`|Especifica que la entidad es una sugerencia de tarea.|
|`URL`| String|`url`|Especifica que la entidad es una dirección URL de Internet.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|Modo de Outlook aplicable| Redacción o lectura|
#### ItemNotificationMessageType :String

Especifica el tipo de mensaje de notificación de una cita o un mensaje.

ItemNotificationMessageType

##### Tipo:

*   String

##### Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`ProgressIndicator`| String|`progressIndicator`|notificationMessage es un indicador de progreso.|
|`InformationalMessage`| String|`informationalMessage`|notificationMessage es un mensaje informativo.|
|`ErrorMessage`| String|`errorMessage`|notificationMessage es un mensaje de error.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|Modo de Outlook aplicable| Redacción o lectura|
#### ItemType :String

Especifica el tipo de un elemento.

ItemType

##### Tipo:

*   String

##### Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`Message`| String|`message`|Un mensaje de correo electrónico o una convocatoria, respuesta o cancelación de una reunión.|
|`Appointment`| String|`appointment`|Un elemento de cita.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|Modo de Outlook aplicable| Redacción o lectura|
#### RecipientType :String

Especifica el tipo de destinatario de una cita.

RecipientType

##### Tipo:

*   String

##### Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`Other`| String|`other`|El destinatario no es uno de los otros tipos de destinatario.|
|`DistributionList`| String|`distributionList`|El destinatario es una lista de distribución que contiene una lista de direcciones de correo electrónico.|
|`User`| String|`user`|El destinatario es una dirección de correo electrónico SMTP que se encuentra en el servidor Exchange.|
|`ExternalUser`| String|`externalUser`|El destinatario es una dirección de correo electrónico SMTP que no se encuentra en el servidor Exchange.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|Modo de Outlook aplicable| Redacción o lectura|
#### ResponseType :String

Especifica el tipo de respuesta para la invitación a una reunión.

ResponseType

##### Tipo:

*   String

##### Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`None`| String|`none`|No ha habido respuesta del asistente.|
|`Organizer`| String|`organizer`|El asistente es el organizador de la reunión.|
|`Tentative`| String|`tentative`|El asistente ha aceptado provisionalmente la convocatoria de reunión.|
|`Accepted`| String|`accepted`|El asistente ha aceptado la convocatoria de reunión.|
|`Declined`| String|`declined`|El asistente ha rechazado la convocatoria de reunión.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1,0|
|Modo de Outlook aplicable| Redacción o lectura|

#### RestVersion :String

Especifica la versión de la API de REST que corresponde a un identificador de elemento con formato REST. 

RestVersion

##### Tipo:

*   String

##### Propiedades:

|Nombre| Tipo| Valor | Descripción|
|---|---|---|---|
|`v1_0`| String|`v1.0`|Versión 1.0.|
|`v2_0`| String|`v2.0`|Versión 2.0.|
|`Beta`| String|`beta`|Beta.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|Modo de Outlook aplicable| Redacción o lectura|
