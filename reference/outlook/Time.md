

# Hora

Se devuelve el objeto `Time` como una propiedad [`start`](Office.context.mailbox.item.md#start-datetime) o [`end`](Office.context.mailbox.item.md#end-datetime) de una cita en modo Redacción.

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción|

### Métodos

####  getAsync([options], callback)

Obtiene la hora de inicio o finalización de una cita.

##### Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult).

La fecha y la hora se proporcionan como un objeto Date en la propiedad `asyncResult.value`. El valor se encuentra en la hora UTC (hora universal coordinada). Puede convertir la hora UTC en la hora local del cliente mediante el método [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción|
####  setAsync(dateTime, [options], [callback])

Establece la hora de inicio o finalización de una cita.

Si se llama al método `setAsync` en la propiedad [`start`](Office.context.mailbox.item.md#start-datetime), la propiedad [`end`](Office.context.mailbox.item.md#end-datetime) se ajustará para mantener la duración de la cita que se ha establecido con anterioridad. Si se llama al método `setAsync` en la propiedad `end`, la duración de la cita se extenderá hasta la nueva hora de finalización.

La hora debe especificarse conforme al sistema UTC. Puede obtener la hora UTC correcta con el método [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date).

##### Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`dateTime`| Fecha||Un objeto Date se encuentra en la hora UTC (hora universal coordinada).|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.<br/><br/>**Propiedades**<br/><table class="nested-table"><thead><tr><th>Nombre</th><th>Tipo</th><th>Atributos</th><th>Descripción</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</td></tr></tbody></table>|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). <br/>Si al establecer la fecha y la hora se produce un error, la propiedad `asyncResult.error` contendrá un código de error.<br/><table class="nested-table"><thead><tr><th>Código de error</th><th>Descripción</th></tr></thead><tbody><tr><td><code>InvalidEndTime</code></td><td>La hora de finalización de la cita se encuentra antes de la hora de inicio de esta.</td></tr></tbody></table>|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](./tutorial-api-requirement-sets.md)| 1.1|
|[Nivel de permisos mínimo](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Modo de Outlook aplicable| Redacción|

##### Ejemplo

En el ejemplo siguiente, se establece la hora de inicio de una cita.

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```
