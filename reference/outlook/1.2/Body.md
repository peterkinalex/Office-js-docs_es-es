

# <a name="body"></a>Cuerpo

El objeto `body` proporciona métodos para agregar y actualizar el contenido del mensaje o la cita. Se devuelve en la propiedad `body` del elemento seleccionado.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

### <a name="methods"></a>Métodos

####  <a name="gettypeasync([options],-[callback])"></a>getTypeAsync([options], [callback])

Obtiene un valor que indica si el contenido tiene formato de texto o HTML.

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

El tipo de contenido se devuelve como uno de los valores de [CoercionType](Office.md#coerciontype-string) de la propiedad `asyncResult.value`.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción|
####  <a name="prependasync(data,-[options],-[callback])"></a>prependAsync(data, [options], [callback])

Agrega el contenido especificado al principio del cuerpo del elemento.

El método `prependAsync` inserta la cadena especificada al principio del cuerpo del elemento. El proceso de llamada al método `prependAsync` es igual que el de llamada al método [`setSelectedDataAsync`](#setselecteddataasyncdata-options-callback), con el punto de inserción al principio del contenido del cuerpo.

Al incluir vínculos en formato HTML, puede deshabilitar la vista previa de vínculo en línea al establecer el atributo `id` en el delimitador (`<a>`) para `LPNoLP`. Por ejemplo:

```
Office.context.mailbox.item.body.prependAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`data`| String||Cadena que se debe insertar al principio del cuerpo. La cadena está limitada a 1 000 000 caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>El formato deseado para el cuerpo. La cadena del parámetro <code>data</code> se convertirá a este formato.</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Cualquier error que se detecte se proporcionará en la propiedad `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>El parámetro <code>data</code> es superior a 1 000 000 de caracteres.</td></tr></tbody></table>|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|
####  <a name="setselecteddataasync(data,-[options],-[callback])"></a>setSelectedDataAsync(data, [options], [callback])

Reemplaza la selección que se ha realizado en el cuerpo por el texto especificado.

El método `setSelectedDataAsync` inserta la cadena especificada en la ubicación del cursor en el cuerpo del elemento o, si el texto se selecciona en el editor, reemplaza el texto seleccionado. Si nuca se ha colocado el cursor en el cuerpo del elemento, o si dicho cuerpo perdió enfoque en la interfaz de usuario, la cadena se insertará en la parte superior del contenido del cuerpo.

Al incluir vínculos en formato HTML, puede deshabilitar la vista previa de vínculo en línea al establecer el atributo `id` en el delimitador (`<a>`) para `LPNoLP`. Por ejemplo:

```
Office.context.mailbox.item.body.setSelectedDataAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`data`| String||Cadena que se debe insertar en el cuerpo. La cadena está limitada a 1 000 000 caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>El formato deseado para el cuerpo. La cadena del parámetro <code>data</code> se convertirá a este formato.</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Cualquier error que se detecte se proporcionará en la propiedad `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>El parámetro <code>data</code> es superior a 1 000 000 de caracteres.</td></tr><tr><td><code>InvalidFormatError</code></td><td>El tipo de cuerpo se establece en HTML y el parámetro de datos contiene texto sin formato.</td></tr></tbody></table>|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|
