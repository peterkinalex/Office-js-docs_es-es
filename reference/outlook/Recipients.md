

# <a name="recipients"></a>Destinatarios

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción|

### <a name="methods"></a>Métodos

####  <a name="addasync(recipients,-[options],-[callback])"></a>addAsync(recipients, [options], [callback])

Agrega una lista de destinatarios a los destinatarios existentes de un mensaje o una cita.

El parámetro `recipients` puede ser una matriz de uno de los elementos siguientes:

*   Cadenas que contienen direcciones de correo electrónico SMTP
*   Objetos `EmailUser`
*   Objetos `EmailAddressDetails`

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`recipients`| Matriz.&lt;(String&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||Destinatarios que se deben agregar a la lista de destinatarios.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Si se produce un error al agregar el destinatario, la propiedad `asyncResult.error` contendrá un código de error.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td>`NumberOfRecipientsExceeded</td><td>El número de destinatarios es superior a 100 entradas.</td></tr></tbody></table>|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|

##### <a name="example"></a>Ejemplo

En el ejemplo siguiente, se crea una matriz de objetos `EmailUser` que se agregan a los destinatarios de la línea Para del mensaje.

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.to.addAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients added");
  }
});
```

####  <a name="getasync([options],-callback)"></a>getAsync([options], callback)

Obtiene una lista de destinatarios para un mensaje o una cita.

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

Cuando finalice la llamada, la propiedad `asyncResult.value` contendrá una matriz de objetos [`EmailAddressDetails`](simple-types.md#emailaddressdetails).|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción|

##### <a name="example"></a>Ejemplo

En el siguiente ejemplo, se obtienen los asistentes opcionales de una reunión.

```js
Office.context.mailbox.item.optionalAttendees.getAsync(function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    var msg = "";
    result.value.forEach(function(recip, index) {
      msg = msg + recip.displayName + " (" + recip.emailAddress + ");";
    });
    showMessage(msg);
  }
});
```

####  <a name="setasync(recipients,-[options],-callback)"></a>setAsync(recipients, [options], callback)

Establece una lista de destinatarios para una cita o un mensaje.

El método `setAsync` sobrescribe la lista de destinatarios actual.

El parámetro `recipients` puede ser una matriz de uno de los elementos siguientes:

*   Cadenas que contienen direcciones de correo electrónico SMTP
*   Objetos `EmailUser`
*   Objetos `EmailAddressDetails`

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`recipients`| Matriz.&lt;(String&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||Destinatarios que se deben agregar a la lista de destinatarios.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Si no es posible establecer el destinatario, la propiedad `asyncResult.error` contendrá un código que indica el error que se produjo al agregar los datos.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td>`NumberOfRecipientsExceeded</td><td>El número de destinatarios es superior a 100 entradas.</td></tr></tbody></table>|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|

##### <a name="example"></a>Ejemplo

En el ejemplo siguiente, se crea una matriz de objetos `EmailUser` y se reemplazan los destinatarios de la línea CC del mensaje por la matriz.

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.cc.setAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients overwritten");
  }
});
```
