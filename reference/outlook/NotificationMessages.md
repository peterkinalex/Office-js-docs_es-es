

# <a name="notificationmessages"></a>NotificationMessages

## <a name="notificationmessages"></a>NotificationMessages

El objeto `NotificationMessages` se devuelve como la propiedad [`notificationMessages`](Office.context.mailbox.item.md#notificationmessages-notificationmessages) de un elemento.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

### <a name="methods"></a>Métodos

####  <a name="addasync(key,-jsonmessage,-[options],-[callback])"></a>addAsync(key, JSONmessage, [options], [callback])

Agrega una notificación a un elemento.

Hay un máximo de 5 notificaciones por mensaje. Establecer más devolverá un error `NumberOfNotificationMessagesExceeded`.

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`key`| String||Una clave especificada por el desarrollador que se usa para hacer referencia a este mensaje de notificación. Los desarrolladores pueden usarla para modificar este mensaje más tarde. No puede tener más de 32 caracteres.|
|`JSONmessage`| Object||Un objeto JSON que contiene el mensaje de notificación que se va a agregar al elemento. Consta de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Descripción</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>Especifica el tipo de mensaje. Si el tipo es <code>ProgressIndicator</code> o <code>ErrorMessage</code>, se suministra un icono automáticamente y el mensaje no es persistente. Por lo tanto, el icono y las propiedades persistentes no son válidos para estos tipos de mensajes. Si se incluyen, se provocará un valor <code>ArgumentException</code>. Si el tipo es <code>ProgressIndicator</code>, el desarrollador debería quitar o reemplazar el indicador de progreso cuando se complete la acción.</td></tr><tr><td><code>icon</code></td><td>String</td><td>Una referencia a un icono que se define en el manifiesto de la sección <code>Resource</code>. Aparece en el área de la barra de información. Solo es aplicable si el tipo es <code>InformationalMessage</code>. Especificar este parámetro para un tipo no admitido produce una excepción.</td></tr><tr><td><code>message</code></td><td>String</td><td>El texto del mensaje de notificación. La longitud máxima es de 150 caracteres. Si el desarrollador pasa en una cadena más larga, se produce una excepción <code>ArgumentOutOfRange</code>.</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>Solo es aplicable cuando el tipo es <code>InformationalMessage</code>. Si es <code>true</code>, el mensaje permanece hasta que el complemento lo quita o el usuario lo descarta. Si es <code>false</code>, se quita cuando el usuario se desplaza a un elemento diferente. Para las notificaciones de error, el mensaje persiste hasta que el usuario lo ve una vez. Si se especifica este parámetro para un tipo no admitido, se produce una excepción.</td></tr></tbody></table>|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). |

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
// Create three notifications, each with a different key
Office.context.mailbox.item.notificationMessages.addAsync("progress", {
  type: "progressIndicator",
  message : "An add-in is processing this message."
});
Office.context.mailbox.item.notificationMessages.addAsync("information", {
  type: "informationalMessage",
  message : "The add-in processed this message.",
  icon : "iconid",
  persistent: false
});
Office.context.mailbox.item.notificationMessages.addAsync("error", {
  type: "errorMessage",
  message : "The add-in failed to process this message."
});
```

####  <a name="getallasync([options],-[callback])"></a>getAllAsync([options], [callback])

Devuelve todas las claves y los mensajes de un elemento.

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

Si se completó correctamente, la propiedad `asyncResult.value` contendrá una matriz de objetos [`NotificationMessageDetails`](simple-types.md#notificationmessagedetails).|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
// Get all notifications
Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult) {
  if (asyncResult.status != "failed") {
    Office.context.mailbox.item.notificationMessages.replaceAsync( "notifications", {
      type: "informationalMessage",
      message : "Found " + asyncResult.value.length + " notifications.",
      icon : "iconid",
      persistent: false
    });
  }
});
```

####  <a name="removeasync(key,-[options],-[callback])"></a>removeAsync(key, [options], [callback])

Quita un mensaje de notificación de un elemento.

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`key`| String||La clave para que se quite el mensaje de notificación.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

Si la clave no se encuentra, se devuelve un error `KeyNotFound` en la propiedad `asyncResult.error`.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
// Remove a notification
Office.context.mailbox.item.notificationMessages.removeAsync("progress");
```

####  <a name="replaceasync(key,-jsonmessage,-[options],-[callback])"></a>replaceAsync(key, JSONmessage, [options], [callback])

Reemplaza un mensaje de notificación que tiene una clave determinada con otro mensaje.

Si no existe un mensaje de notificación con la clave especificada, `replaceAsync` agregará la notificación.

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`key`| String||La clave para que se reemplace el mensaje de notificación. No puede ser superior a 32 caracteres.|
|`JSONmessage`| Object||Un objeto JSON que contiene el nuevo mensaje de notificación para reemplazar al mensaje existente. Consta de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Descripción</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>Especifica el tipo de mensaje. Si el tipo es <code>ProgressIndicator</code> o <code>ErrorMessage</code>, se suministra un icono automáticamente y el mensaje no es persistente. Por lo tanto, el icono y las propiedades persistentes no son válidos para estos tipos de mensajes. Si se incluyen, se provocará un valor <code>ArgumentException</code>. Si el tipo es <code>ProgressIndicator</code>, el desarrollador debería quitar o reemplazar el indicador de progreso cuando se complete la acción.</td></tr><tr><td><code>icon</code></td><td>String</td><td>Una referencia a un icono que se define en el manifiesto de la sección <code>Resource</code>. Aparece en el área de la barra de información. Solo es aplicable si el tipo es <code>InformationalMessage</code>. Especificar este parámetro para un tipo no admitido produce una excepción.</td></tr><tr><td><code>message</code></td><td>String</td><td>El texto del mensaje de notificación. La longitud máxima es de 150 caracteres. Si el desarrollador pasa en una cadena más larga, se produce una excepción <code>ArgumentOutOfRange</code>.</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>Solo es aplicable cuando el tipo es <code>InformationalMessage</code>. Si es <code>true</code>, el mensaje permanece hasta que el complemento lo quita o el usuario lo descarta. Si es <code>false</code>, se quita cuando el usuario se desplaza a un elemento diferente. Para las notificaciones de error, el mensaje persiste hasta que el usuario lo ve una vez. Si se especifica este parámetro para un tipo no admitido, se produce una excepción.</td></tr></tbody></table>|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). |

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
// Replace a notification with an informational notification
Office.context.mailbox.item.notificationMessages.replaceAsync("progress", {
  type: "informationalMessage",
  message : "The message was processed successfully.",
  icon : "iconid",
  persistent: false
});
```
