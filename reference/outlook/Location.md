

# <a name="location"></a>Location

Proporciona métodos para obtener y establecer la ubicación de una cita en un complemento de Outlook.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción|

### <a name="methods"></a>Métodos

####  <a name="getasync([options],-callback)"></a>getAsync([options], callback)

Obtiene la ubicación de una cita.

El método `getAsync` inicia una llamada asincrónica al servidor de Exchange para obtener la ubicación de una cita.

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

La ubicación de la cita se proporciona como una cadena en la propiedad `asyncResult.value`.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción|
####  <a name="setasync(location,-[options],-[callback])"></a>setAsync(location, [options], [callback])

Establece la ubicación de una cita.

El método `setAsync` inicia una llamada asincrónica al servidor de Exchange para establecer la ubicación de una cita. Cuando se establezca, se sobrescribirá la ubicación actual.

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`location`| String||La ubicación de la cita. La cadena está limitada a 255 caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Si se produce un error en el establecimiento de la ubicación, la propiedad `asyncResult.error` contendrá un código de error.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>El parámetro <code>location</code> es superior a 255 caracteres.</td></tr></tbody></table>|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción|
