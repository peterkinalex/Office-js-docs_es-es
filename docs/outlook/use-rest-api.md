# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a>Usar las API de REST de Outlook desde un complemento de Outlook

El espacio de nombres [Office.context.mailbox.item](..\..\reference\outlook\Office.context.mailbox.item.md) proporciona acceso a muchos de los campos comunes de mensajes y citas. En cambio, en algunos escenarios, un complemento puede necesitar tener acceso a los datos que no expone el espacio de nombres. Por ejemplo, el complemento puede basarse en propiedades personalizadas que establece una aplicación externa, o necesita buscar el buzón del usuario para obtener mensajes del mismo remitente. En estos escenarios, las [API de REST de Outlook](https://dev.outlook.com/restapi/reference) es el método recomendado para recuperar la información.

## <a name="get-an-access-token"></a>Obtener un token de acceso

Las API de REST de Outlook necesitan un token de portador en el encabezado `Authorization`. Normalmente, las aplicaciones usan flujos de OAuth2 para recuperar un token. En cambio, los complementos pueden recuperar un token sin implementar OAuth2 con el nuevo método [Office.context.mailbox.getCallbackTokenAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#getCallbackTokenAsync) que se ha presentado en la versión preliminar 1.5 del conjunto de requisitos del buzón.

> **Nota**: Como el conjunto de requisitos del buzón 1.5 está en versión preliminar, no puede especificarlo como un requisito en el manifiesto. 

Al establecer la opción `isRest` en `true`, puede solicitar un token compatible con las API de REST.

### <a name="add-in-permissions-and-token-scope"></a>Agregar permisos y ámbito del token

Es importante tener en cuenta el nivel de acceso que necesitará el complemento mediante las API de REST. En la mayoría de los casos, el token que se ha devuelto mediante `getCallbackTokenAsync`, proporcionará acceso de solo lectura únicamente al elemento actual. Esto se cumple incluso si el complemento especifica el nivel de permiso `ReadWriteItem` en su manifiesto.

Si el complemento va a necesitar acceso de escritura para el elemento actual o para otros elementos del buzón del usuario, su complemento debe especificar el nivel de permiso `ReadWriteMailbox` en su manifiesto. En este caso, el token devuelto contendrá acceso de lectura y escritura a los contactos, eventos y mensajes del usuario.

### <a name="example"></a>Ejemplo

```js
Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    var accessToken = result.value;
    
    // Use the access token
    getCurrentItem(accessToken);
  } else {
    // Handle the error
  }
});
```

## <a name="get-the-item-id"></a>Obtener el identificador de elemento

Para recuperar el elemento actual mediante REST, su complemento necesitará el id. de elemento, con un formato correcto para REST. Este se obtiene de la propiedad [Office.context.mailbox.item.itemId](../../reference/outlook/Office.context.mailbox.item.md), pero deben realizarse algunas comprobaciones para asegurarse de que es un id. con formato para REST.

- En Outlook Mobile, el valor devuelto por `Office.context.mailbox.item.itemId` es un id. con formato para REST y puede usarse como está.
- En otros clientes de Outlook, el valor devuelto por `Office.context.mailbox.item.itemId` es un id. con formato para EWS, y debe convertirse con el método [Office.context.mailbox.convertToRestId](../../reference/outlook/Office.context.mailbox.md).

El complemento puede determinar qué cliente de Outlook está cargado mediante la comprobación de la propiedad [Office.context.mailbox.diagnostics.hostName](../../reference/outlook/Office.context.mailbox.diagnostics.md).

### <a name="example"></a>Ejemplo

```js
function getItemRestId() {
  // Currently the only Outlook Mobile version that supports add-ins
  // is Outlook for iOS.
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}
```

## <a name="get-the-rest-api-url"></a>Obtener la dirección URL de la API de REST

La última información que el complemento necesita para llamar a la API de REST es el nombre de host que debe usar para enviar solicitudes de API. Esta información se encuentra en la propiedad [Office.context.mailbox.restUrl](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#restUrl).

### <a name="example"></a>Ejemplo

```js
// Example: https://outlook.office.com
var restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a>Llamar a la API

Una vez que el complemento tiene el token de acceso, el id. de elemento y la dirección URL de la API de REST, puede pasar esa información a un servicio de back-end que llame a la API de REST o puede llamarla directamente con AJAX. En el ejemplo siguiente se llama a la API de REST de Correo de Outlook para obtener el mensaje actual.

```js
function getCurrentItem(accessToken) {
  // Get the item's REST ID
  var itemId = getItemRestId();

  // Construct the REST URL to the current item
  // Details for formatting the URL can be found at 
  // https://msdn.microsoft.com/office/office365/APi/mail-rest-operations#get-a-message-rest
  var getMessageUrl = Office.context.mailbox.restUrl +
    '/api/v2.0/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`
    var subject = item.Subject;
    ...
  }).fail(function(error){
    // Handle error
  });
}
```

## <a name="additional-resources"></a>Recursos adicionales

Para obtener un ejemplo que llame a las API de REST desde un complemento de Outlook, vea [command-demo](https://github.com/jasonjoh/command-demo) en GitHub.