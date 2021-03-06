

# <a name="mailbox"></a>buzón de correo

## [Office](Office.md)[.context](Office.context.md). mailbox

Proporciona acceso al modelo de objetos del complemento de Outlook para Microsoft Outlook y Microsoft Outlook en la web.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|

### <a name="namespaces"></a>Espacios de nombres

[diagnostics](Office.context.mailbox.diagnostics.md): Proporciona información de diagnóstico a un complemento de Outlook.

[item](Office.context.mailbox.item.md): Proporciona métodos y propiedades para tener acceso a un mensaje o cita en un complemento de Outlook.

[userProfile](Office.context.mailbox.userProfile.md): Proporciona información sobre el usuario en un complemento de Outlook.</dd>

### <a name="members"></a>Miembros

#### <a name="ewsurl-string"></a>ewsUrl :String

Obtiene la URL del punto de conexión de los Servicios web de Exchange (EWS) para esta cuenta de correo electrónico. Solo modo Lectura.

> **Nota:** Este miembro no se admite en Outlook para iOS ni en Outlook para Android.

El valor `ewsUrl` puede usarse por un servicio remoto para realizar llamadas EWS al buzón del usuario. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx).

La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al miembro `ewsUrl` en el modo lectura.

En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item#saveAsync) antes de poder usar el miembro `ewsUrl`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

### <a name="methods"></a>Métodos

####  <a name="converttoewsiditemid-restversion--string"></a>convertToEwsId(itemId, restVersion) → {String}

Convierte un identificador de elemento con formato para REST al formato EWS.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

Los identificadores de elemento obtenidos a través de una API de REST (como la [API de correo de Outlook](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations) o [Microsoft Graph](http://graph.microsoft.io/)) usan un formato diferente al formato que usa Exchange Web Services (EWS). El método `convertToEwsId` convierte un identificador con formato REST al formato adecuado para EWS.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`itemId`| String|Un identificador de elemento con formato para las API de REST de Outlook|
|`restVersion`| [Office.MailboxEnums.RestVersion](Office.MailboxEnums.md#restversion)|Un valor que indica la versión de la API de REST de Outlook que se usa para recuperar el identificador de elemento.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="returns"></a>Valores devueltos:

Tipo: String

##### <a name="example"></a>Ejemplo

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimesimple-typesmdlocalclienttime"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](simple-types.md#localclienttime)}

Obtiene un diccionario con información de tiempo en el tiempo del cliente local.

Las fechas y horas usadas por una aplicación de correo para Outlook o Outlook Web App pueden usar distintas zonas horarias. Outlook usa la zona horaria del equipo cliente; Outlook Web App usa la zona horaria definida en el Centro de administración de Exchange (EAC). Debería tratar los valores de fecha y hora de modo que los valores que aparezcan en la interfaz de usuario sean siempre coherentes con la zona horaria que el usuario espera.

Si se está ejecutando la aplicación de correo en Outlook, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria del equipo cliente. Si se está ejecutando la aplicación de correo en Outlook Web App, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria especificada en el CEF.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`timeValue`| Fecha|Un objeto Date|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="returns"></a>Valores devueltos:

Tipo: [LocalClientTime](simple-types.md#localclienttime)

####  <a name="converttorestiditemid-restversion--string"></a>convertToRestId(itemId, restVersion) → {String}

Convierte un identificador de elemento con formato para EWS al formato REST.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

Los identificadores de elemento obtenidos a través de EWS o de la propiedad `itemId` usan un formato diferente al formato que usan las API de REST (como la [API de correo de Outlook](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations) o [Microsoft Graph](http://graph.microsoft.io/)). El método `convertToRestId` convierte un identificador con formato EWS al formato adecuado para REST.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`itemId`| String|Un identificador de elemento con formato para Exchange Web Services (EWS)|
|`restVersion`| [Office.MailboxEnums.RestVersion](Office.MailboxEnums.md#restversion)|Un valor que indica la versión de la API de REST de Outlook con la que se usará el identificador convertido.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="returns"></a>Valores devueltos:

Tipo: String

##### <a name="example"></a>Ejemplo

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

Obtiene un objeto Date del diccionario que contiene información de tiempo.

El método `convertToUtcClientTime` convierte un diccionario que contiene la fecha y la hora locales en un objeto Date con los valores correctos para la fecha y la hora locales.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`input`| [LocalClientTime](simple-types.md#localclienttime)|El valor de la hora local para convertir.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="returns"></a>Valores devueltos:

Objeto Date con el tiempo expresado en UTC.

<dl class="param-type">

<dt>Tipo</dt>

<dd>Fecha</dd>

</dl>

####  <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

Muestra una cita de calendario existente.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

El método `displayAppointmentForm` abre una cita de calendario existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.

En Outlook para Mac, puede usar este método para mostrar una cita que no forme parte de una serie periódica o la cita principal de una serie periódica, pero no puede mostrar una instancia de la serie. Esto es porque en Outlook para Mac no se puede tener acceso a las propiedades (incluido el identificador de elemento) de las instancias de una serie periódica.

En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.

Si el identificador de elemento especificado no identifica una cita existente, se abrirá una página en blanco en el dispositivo o equipo cliente y no se generará ningún mensaje de error.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`itemId`| String|Identificador de los servicios web de Exchange (EWS) para una cita de calendario existente.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

Muestra un mensaje existente.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

El método `displayMessageForm` abre un mensaje existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.

En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.

Si el identificador de elemento especificado no identifica un mensaje existente, no se mostrará ningún mensaje en el equipo cliente ni se generará ningún mensaje de error.

No use el valor `displayMessageForm` con un `itemId` que represente una cita. Use el método `displayAppointmentForm` para mostrar una cita existente y `displayNewAppointmentForm` para mostrar un formulario para crear una cita nueva.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`itemId`| String|Identificador de los servicios web de Exchange (EWS) para un mensaje existente.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

Muestra un formulario para crear una nueva cita de calendario.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

El método `displayNewAppointmentForm` abre un formulario que permite al usuario crear una nueva cita o reunión. Si se especifican parámetros, los campos de formulario de cita se rellenan automáticamente con el contenido de los parámetros.

En Outlook Web App y OWA para dispositivos, este método muestra siempre un formulario con un campo de asistentes. Si no especifica ningún asistente como argumento de entrada, el método muestra un formulario con un botón **Guardar**. Si ha especificado asistentes, el formulario incluirá a los asistentes y un botón **Enviar**.

En el cliente enriquecido de Outlook y Outlook RT, si se especifica cualquier asistente o recurso en los parámetros `requiredAttendees`, `optionalAttendees` o `resources`, este método muestra un formulario de reunión con un botón **Enviar**. Si no se especifica ningún destinatario, este método muestra un formulario de cita con un botón **Guardar y cerrar**.

Si cualquiera de los parámetros supera los límites de tamaño especificados o si se especifica un nombre de parámetro desconocido, se genera una excepción.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`parameters`| Object|Un diccionario de parámetros que describen la nueva cita.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Descripción</th></tr></thead><tbody><tr><td><code>requiredAttendees</code></td><td>Array.&lt;String&gt; &#124; Array.&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>Una matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto <code>EmailAddressDetails</code> para cada uno de los asistentes necesarios de la cita. La matriz está limitada a un máximo de 100 entradas.</td></tr><tr><td><code>optionalAttendees</code></td><td>Array.&lt;String&gt; &#124; Array.&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>Matriz de cadenas que contiene las direcciones de correo o matriz que contiene un EmailAddressDetails object por cada uno de los asistentes opcionales de la cita. La matriz está limitada a un máximo de 100 entradas.</td></tr><tr><td><code>start</code></td><td>Fecha</td><td>Un objeto Date que especifica la fecha y hora de inicio de la cita.</td></tr><tr><td><code>end</code></td><td>Fecha</td><td>Un objeto Date que especifica la fecha y hora de finalización de la cita.</td></tr><tr><td><code>location</code></td><td>String</td><td>Una cadena que contiene la ubicación de la cita. La cadena está limitada a un máximo de 255 caracteres.</td></tr><tr><td><code>resources</code></td><td>Array.&lt;String&gt;</td><td>Una matriz de cadenas que contiene los recursos necesarios para la cita. La matriz está limitada a un máximo de 100 entradas.</td></tr><tr><td><code>subject</code></td><td>String</td><td>Una cadena que contiene al asunto de la cita. La cadena está limitada a un máximo de 255 caracteres.</td></tr><tr><td><code>body</code></td><td>String</td><td>El cuerpo del mensaje de la cita. El contenido del cuerpo está limitado a un tamaño máximo de 32 KB.</td></tr></tbody></table>|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Lectura|

##### <a name="example"></a>Ejemplo

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

Obtiene una cadena que contiene un token usado para obtener datos adjuntos o un elemento de Exchange Server.

El método `getCallbackTokenAsync` realiza una llamada asincrónica para obtener un token opaco desde Exchange Server que hospeda el buzón del usuario. La duración del token de devolución de llamada es de 5 minutos.

Puede pasar el token y un identificador de archivo adjunto o el identificador del elemento a un sistema de terceros. El sistema de terceros usa el token como token de autorización del portador para llamar a los Servicios web de Exchange (EWS) o a las operaciones [GetAttachment](https://msdn.microsoft.com/en-us/library/office/aa494316.aspx) o [GetItem](https://msdn.microsoft.com/en-us/library/office/aa565934.aspx) para devolver datos adjuntos o un elemento. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx).

La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al método `getCallbackTokenAsync` en el modo lectura.

En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item#saveAsync) para obtener un identificador de elemento para pasar al método `getCallbackTokenAsync`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). El token se proporciona como una cadena en la propiedad `asyncResult.value`.|
|`userContext`| Objeto| &lt;optional&gt;|Cualquier dato de estado que se pasa al método asincrónico.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción y lectura|

##### <a name="example"></a>Ejemplo

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync(callback, [userContext])

Obtiene un token que identifica al usuario y al complemento de Office.

El método `getUserIdentityTokenAsync` devuelve un token que puede usar para identificar y [autenticar el complemento y el usuario mediante un sistema de terceros](https://msdn.microsoft.com/EN-US/library/office/fp179828.aspx).

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

El token se proporciona como una cadena en la propiedad `asyncResult.value`. | |`userContext`| Object | &lt;optional&gt;| Los datos de estado que se pasan al método asincrónico. |

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync(data, callback, [userContext])

Realiza una solicitud asincrónica a un servicio de Servicios Web Exchange (EWS) en el servidor Exchange que hospeda el buzón del usuario.

> **Nota:** Este método no se admite en Outlook para iOS ni en Outlook para Android.

El método `makeEwsRequestAsync` envía una solicitud de EWS en nombre del complemento a Exchange.

No puede solicitar elementos asociados de las carpetas con el método `makeEwsRequestAsync`.

La solicitud XML debe especificar codificación UTF-8.

```
<?xml version="1.0" encoding="utf-8"?>
```

Su complemento debe tener el permiso **ReadWriteMailbox** para usar el método `makeEwsRequestAsync`. Para obtener información sobre cómo usar el permiso **ReadWriteMailbox** y sobre las operaciones EWS que puede llamar con el método `makeEwsRequestAsync`, consulte [Especificar permisos para el acceso del complemento de correo al buzón del usuario](../../docs/outlook/understanding-outlook-add-in-permissions.md).

**NOTA**: El administrador del servidor debe establecer `OAuthAuthentication` como True en el directorio EWS del servidor de acceso de cliente para permitir que el método `makeEwsRequestAsync` realice las solicitudes de EWS.

#### <a name="version-differences"></a>Diferencias de versión

Si usa el método `makeEwsRequestAsync` en aplicaciones de correo que se ejecutan en versiones de Outlook anteriores a 15.0.4535.1004, debe establecer el valor de codificación a `ISO-8859-1`.

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

No es necesario establecer el valor de codificación si la aplicación de correo se ejecuta en Outlook en la web. Puede averiguar si su aplicación de correo se ejecuta en Outlook o en Outlook en la web usando la propiedad mailbox.diagnostics.hostName. Para averiguar qué versión de Outlook se está ejecutando, use la propiedad mailbox.diagnostics.hostVersion.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`data`| String||La solicitud de EWS.|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

El resultado XML de la llamada EWS se proporciona como una cadena en la propiedad `asyncResult.value`. Si el resultado supera 1 MB de tamaño, se devuelve un mensaje de error en su lugar. | |`userContext`| Objeto | &lt;opcional&gt;| Los datos de estado que se pasan al método asincrónico.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteMailbox|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

En el siguiente ejemplo, se llama a `makeEwsRequestAsync` para usar la operación `GetItem` para obtener el asunto de un elemento.

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```
