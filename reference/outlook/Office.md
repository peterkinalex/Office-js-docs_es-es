 

# <a name="office"></a>Office

El espacio de nombres de Office proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres de Office, vea [API compartida](../shared/shared-api.md).

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción o lectura|

### <a name="namespaces"></a>Espacios de nombres

[context](Office.context.md): Proporciona interfaces compartidas del espacio de nombres de contexto de la API de complementos de Office para su uso en la API de complemento de Outlook.

[MailboxEnums](Office.MailboxEnums.md): Incluye las enumeraciones ItemType, EntityType, AttachmentType, RecipientType, ResponseType y ItemNotificationMessageType.

### <a name="members"></a>Miembros

####  <a name="asyncresultstatus-:string"></a>AsyncResultStatus :String

Especifica el resultado de una llamada asíncrona.

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`Succeeded`| String|La llamada ha sido correcta.|
|`Failed`| String|La llamada ha fallado.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción o lectura|
####  <a name="coerciontype-:string"></a>CoercionType :String

Especifica cómo convertir los datos que el método invocado ha devuelto o definido.

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`Html`| String|Solicita que los datos se devuelvan en formato HTML.|
|`Text`| String|Solicita que los datos se devuelvan en formato de texto.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción o lectura|
####  <a name="sourceproperty-:string"></a>SourceProperty :String

Especifica el origen de los datos devueltos por el método invocado.

##### <a name="type:"></a>Tipo:

*   String

##### <a name="properties:"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`Body`| String|El origen de los datos proviene del cuerpo de un mensaje.|
|`Subject`| String|El origen de los datos proviene del asunto de un mensaje.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|Modo de Outlook aplicable| Redacción o lectura|
