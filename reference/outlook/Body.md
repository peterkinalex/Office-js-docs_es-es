

# Cuerpo

El objeto `body` proporciona métodos para agregar y actualizar el contenido del mensaje o la cita. Se devuelve en la propiedad `body` del elemento seleccionado.

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

### Métodos

####  getAsync(coercionType, [options], [callback])

Devuelve el cuerpo actual en un formato especificado.

Este método devuelve todo el cuerpo actual en el formato especificado por `coercionType`.

##### Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||El formato del cuerpo devuelto.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

El cuerpo se proporciona en el formato solicitado en la propiedad `asyncResult.value`.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### Ejemplos

Este ejemplo obtiene el cuerpo del mensaje en texto sin formato.

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Do something with the result
  });
```

A continuación, se muestra un ejemplo del parámetro `result` que se ha pasado a la función de devolución de llamada.

```js
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  getTypeAsync([options], [callback])

Obtiene un valor que indica si el contenido tiene formato de texto o HTML.

##### Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

El tipo de contenido se devuelve como uno de los valores de [CoercionType](Office.md#coerciontype-string) de la propiedad `asyncResult.value`.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción|
####  prependAsync(data, [options], [callback])

Agrega el contenido especificado al principio del cuerpo del elemento.

El método `prependAsync` inserta la cadena especificada al principio del cuerpo del elemento. El proceso de llamada al método `prependAsync` es igual que el de llamada al método [`setSelectedDataAsync`](#setselecteddataasync), con el punto de inserción al principio del contenido del cuerpo.

Al incluir vínculos en formato HTML, puede deshabilitar la vista previa de vínculo en línea al establecer el atributo `id` en el delimitador (`<a>`) para `LPNoLP`. Por ejemplo:

```js
Office.context.mailbox.item.body.prependAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`data`| String||Cadena que se debe insertar al principio del cuerpo. La cadena está limitada a 1 000 000 caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>El formato deseado para el cuerpo. La cadena del parámetro <code>data</code> se convertirá a este formato.</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Cualquier error que se detecte se proporcionará en la propiedad `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>El parámetro <code>data</code> es superior a 1 000 000 de caracteres.</td></tr></tbody></table>|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|
####  setAsync(data, [options], [callback])

Reemplaza todo el cuerpo con el texto especificado.

El método `setAsync` reemplaza el cuerpo del elemento existente con la cadena especificada o, si se selecciona texto en el editor, reemplaza el texto seleccionado.

Al incluir vínculos en formato HTML, puede deshabilitar la vista previa de vínculo en línea al establecer el atributo `id` en el delimitador (`<a>`) para `LPNoLP`. Por ejemplo:

```js
Office.context.mailbox.item.body.setAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`data`| String||La cadena que reemplazará el cuerpo existente. La cadena está limitada a 1 000 000 de caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>El formato deseado para el cuerpo. La cadena del parámetro <code>data</code> se convertirá a este formato.</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Cualquier error que se detecte se proporcionará en la propiedad `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>El parámetro <code>data</code> es superior a 1 000 000 de caracteres.</td></tr><tr><td><code>InvalidFormatError</code></td><td>El parámetro <code>options.coercionType</code> se establece en <code>Office.CoercionType.Html</code> y el cuerpo del mensaje está en texto sin formato.</td></tr></tbody></table>|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.3|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|

##### Ejemplos

En el ejemplo siguiente, se reemplaza el cuerpo con contenido HTML.

```js
Office.context.mailbox.item.body.setAsync(
  "<b>(replaces all body, including threads you are replying to that may be on the bottom)</b>",
  { coercionType:"html", asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Process the result
  });
```

A continuación, se muestra un ejemplo del parámetro `result` que se ha pasado a la función de devolución de llamada.

```js
{
  "value":null,
  "status":"succeeded",
  "asyncContext":"This is passed to the callback"
}
```

####  setSelectedDataAsync(data, [options], [callback])

Reemplaza la selección que se ha realizado en el cuerpo por el texto especificado.

El método `setSelectedDataAsync` inserta la cadena especificada en la ubicación del cursor en el cuerpo del elemento o, si el texto se selecciona en el editor, reemplaza el texto seleccionado. Si nuca se ha colocado el cursor en el cuerpo del elemento, o si dicho cuerpo perdió enfoque en la interfaz de usuario, la cadena se insertará en la parte superior del contenido del cuerpo.

Al incluir vínculos en formato HTML, puede deshabilitar la vista previa de vínculo en línea al establecer el atributo `id` en el delimitador (`<a>`) para `LPNoLP`. Por ejemplo:

```js
Office.context.mailbox.item.body.setSelectedDataAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`data`| String||Cadena que se debe insertar en el cuerpo. La cadena está limitada a 1 000 000 caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>El formato deseado para el cuerpo. La cadena del parámetro <code>data</code> se convertirá a este formato.</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Cualquier error que se detecte se proporcionará en la propiedad `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>El parámetro <code>data</code> es superior a 1 000 000 de caracteres.</td></tr><tr><td><code>InvalidFormatError</code></td><td>El tipo de cuerpo se establece en HTML y el parámetro de datos contiene texto sin formato.</td></tr></tbody></table>|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|
