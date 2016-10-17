 

# <a name="mailboxenums"></a>MailboxEnums

## [Office](Office.md). MailboxEnums

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción|

### <a name="members"></a>Miembros

#### <a name="attachmenttype-:string"></a>AttachmentType :String

Especifica el tipo de datos adjuntos. Solo  modo Redacción.

AttachmentType

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`File`| String|Los datos adjuntos son un archivo.|
|`Item`| String|Los datos adjuntos son un elemento de Exchange.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción|
#### <a name="entitytype-:string"></a>EntityType :String

Especifica un tipo de entidad. Solo  modo Redacción.

EntityType

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`Address`| String|Especifica que la entidad es una dirección postal.|
|`Contact`| String|Especifica que la entidad es un contacto.|
|`EmailAddress`| String|Especifica que la entidad es una dirección de correo electrónico SMTP.|
|`MeetingSuggestion`| String|Especifica que la entidad es una sugerencia de reunión.|
|`PhoneNumber`| String|Especifica que la entidad es un número de teléfono de EE. UU.|
|`TaskSuggestion`| String|Especifica que la entidad es una sugerencia de tarea.|
|`URL`| String|Especifica que la entidad es una dirección URL de Internet.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción|
#### <a name="itemtype-:string"></a>ItemType :String

Especifica el tipo de un elemento. Solo  modo Redacción.

ItemType

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`Message`| String|Un mensaje de correo electrónico o una convocatoria, respuesta o cancelación de una reunión.|
|`Appoinment`| String|Un elemento de cita.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción|
#### <a name="recipienttype-:string"></a>RecipientType :String

Especifica el tipo de destinatario de una cita. Solo en modo Redacción.

RecipientType

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`Other`| String|El destinatario no es uno de los otros tipos de destinatario.|
|`DistributionList`| String|El destinatario es una lista de distribución que contiene una lista de direcciones de correo electrónico.|
|`User`| String|El destinatario es una dirección de correo electrónico SMTP que se encuentra en el servidor Exchange.|
|`ExternalUser`| String|El destinatario es una dirección de correo electrónico SMTP que no se encuentra en el servidor Exchange.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.1|
|Modo de Outlook aplicable| Redacción|
#### <a name="responsetype-:string"></a>ResponseType :String

Especifica el tipo de respuesta para la invitación a una reunión. Solo modo Redacción.

ResponseType

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`None`| String|No ha habido respuesta del asistente.|
|`Organizer`| String|El asistente es el organizador de la reunión.|
|`Tentative`| String|El asistente ha aceptado provisionalmente la convocatoria de reunión.|
|`Accepted`| String|El asistente ha aceptado la convocatoria de reunión.|
|`Declined`| String|El asistente ha rechazado la convocatoria de reunión.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción|
