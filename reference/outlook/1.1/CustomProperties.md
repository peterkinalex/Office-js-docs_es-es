

# CustomProperties

El objeto `CustomProperties` representa las propiedades personalizadas que son específicas de un elemento determinado y de un complemento de correo de Outlook concreto. Por ejemplo, puede ser necesario que un complemento de correo guarde algunos datos que son específicos del mensaje de correo electrónico actual que ha activado el complemento. Si el usuario vuelve a visitar el mismo mensaje y activa el complemento de correo de nuevo, el complemento será capaz de recuperar los datos que se han guardado como propiedades personalizadas.

Debido a que Outlook para Mac no almacena propiedades personalizadas en la memoria caché, si se desconecta la red del usuario, los complementos de correo no podrán tener acceso a sus propiedades personalizadas.

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

### Ejemplo

En el ejemplo siguiente se muestra cómo usar el método `loadCustomPropertiesAsync` para cargar de forma asincrónica propiedades personalizadas que son específicas del elemento actual. En el ejemplo también se muestra cómo usar el método [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, se usa el método [`get`](CustomProperties.md#get) para leer la propiedad personalizada `myProp`, el método [`set`](CustomProperties.md#set) para escribir en la propiedad personalizada `otherProp`, y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.

```
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var mailbox = Office.context.mailbox;
    mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
  });
}
function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

### Métodos

####  get(name) → {String}

Devuelve el valor de la propiedad personalizada especificada.

##### Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`name`| String|Nombre de la propiedad personalizada que se devolverá.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### Valores devueltos:

Valor de la propiedad personalizada especificada.

<dl class="param-type">

<dt>Tipo</dt>

<dd>String</dd>

</dl>

####  remove(name)

Elimina la propiedad especificada de la colección de propiedades personalizadas.

Para que la eliminación de la propiedad sea permanente, debe llamar al método [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) del objeto `CustomProperties`.

##### Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`name`| String|Nombre de la propiedad que se va a eliminar.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|
####  saveAsync([callback], [asyncContext])

Guarda propiedades personalizadas específicas del elemento en el servidor.

Debe llamar al método `saveAsync` para almacenar cualquier cambio realizado con el método [`set`](CustomProperties.md#set) o el método [`remove`](CustomProperties.md#remove) del objeto `CustomProperties`. La acción de guardado es asincrónica.

Una buena práctica es hacer que la función de devolución de llamada compruebe y controle errores de `saveAsync`. En particular, un complemento de lectura puede activarse mientras el usuario está en un estado de conexión en un formulario de lectura y, posteriormente, se desconecta al usuario. Si el complemento llama a `saveAsync` en el estado desconectado, `saveAsync` devolverá un error. El método de devolución de llamada debe controlar este error de forma adecuada.

##### Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). |
|`asyncContext`| Object| &lt;optional&gt;|Cualquier dato de estado que se pasa al método de devolución de llamada.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|

##### Ejemplo

En el siguiente ejemplo de código de JavaScript se muestra cómo usar de forma asincrónica el método `loadCustomPropertiesAsync` para cargar propiedades personalizadas específicas del elemento actual, y el método [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, el ejemplo de código usa el método [`get`](CustomProperties.md#get) para leer la propiedad personalizada `myProp`, el método [`set`](CustomProperties.md#set) para escribir en la propiedad personalizada `otherProp`, y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
}
function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed){
    write(asyncResult.error.message);
  }
  else {
    // Async call to save custom properties completed.
    // Proceed to do the appropriate for your add-in.
  }
}

// Writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  set(name, value)

Configura la propiedad especificada al valor especificado.

El método `set` configura la propiedad especificada al valor especificado. Debe usar el método [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) para guardar la propiedad en el servidor.

El método `set` crea una propiedad nueva si la propiedad especificada aún no existe. De lo contrario, el valor existente se sustituye por el nuevo valor. El parámetro `value` puede ser de cualquier tipo, pero siempre se pasa al servidor como una cadena.

##### Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`name`| String|Nombre de la propiedad que se va a configurar.|
|`value`| Object|Valor de la propiedad que se va a configurar.|

##### Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1,0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Modo de Outlook aplicable| Redacción o lectura|