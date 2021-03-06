

# <a name="roamingsettings"></a>RoamingSettings

La configuración creada mediante los métodos del objeto `RoamingSettings` se guarda por complemento y por usuario. Es decir, está disponible solo para el complemento que la ha creado y solo para el buzón de correo del usuario en el que se guarda.

> Aunque la API del complemento de Outlook limita el acceso a esta configuración solo al complemento que la creó, esta configuración no se debe considerar un almacenamiento seguro. Se puede tener acceso mediante Servicios Web Exchange o la biblioteca MAPI extendida. No debe usarse para almacenar información confidencial como credenciales de usuario o tokens de seguridad.

El nombre de una configuración es una cadena, mientras que el valor puede ser String, Number, Boolean, Null, Object o Array.

Se puede tener acceso al objeto `RoamingSettings` mediante la propiedad [`roamingSettings`](Office.context.md#roamingsettings-roamingsettings) del espacio de nombres `Office.context`.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|

### <a name="example"></a>Ejemplo

```JavaScript
// Get the current value of the 'myKey' setting
var value = Office.context.roamingSettings.get('myKey');
// Update the value of the 'myKey' setting
Office.context.roamingSettings.set('myKey', 'Hello World!');
// Persist the change
Office.context.roamingSettings.saveAsync();
```

### <a name="methods"></a>Métodos

####  <a name="get(name)-→-(nullable)-{string|number|boolean|object|array}"></a>get(name) → (nullable) {String|Number|Boolean|Object|Array}

Recupera la configuración especificada.

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`name`| String|El nombre con distinción de mayúsculas y minúsculas de la configuración que se debe recuperar.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|

##### <a name="returns:"></a>Valores devueltos:

<dl class="param-type">

<dt>Type</dt>

<dd>String | Number | Boolean | Object | Array</dd>

</dl>

####  <a name="remove(name)"></a>remove(name)

Elimina la configuración especificada.

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`name`| String|El nombre con distinción de mayúsculas y minúsculas de la configuración que se debe eliminar.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|
####  <a name="saveasync([callback])"></a>saveAsync([callback])

Guarda la configuración.

Al inicializar un complemento, se cargarán todas las configuraciones que haya guardado. Esto significa que, durante la sesión, solo podrá usar los métodos [`set`](RoamingSettings.md#set) y [`get`](RoamingSettings.md#get) para trabajar con la copia en memoria del contenedor de propiedades de configuración. Si quiere guardar la configuración para que esté disponible la próxima vez que use el complemento, use el método `saveAsync`.

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](simple-types.md#asyncresult). |

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|
####  <a name="set(name,-value)"></a>set(name, value)

Define o crea la configuración especificada.

El método set crea una nueva configuración del nombre especificado si no existe todavía o establece para él mismo una configuración ya existente. El valor se almacena en el documento como la representación JSON serializada del tipo de datos correspondiente.

El espacio máximo disponible para la configuración de cada complemento es de 2 MB y cada configuración individual está limitada a 32 KB.

Los cambios realizados a la configuración mediante la función `set` no se guardarán en el servidor hasta que se llame a la función [`saveAsync`](RoamingSettings.md#saveasynccallback).

##### <a name="parameters:"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`name`| String|Nombre, con distinción entre mayúsculas y minúsculas, de la configuración que se debe establecer o crear.|
|`value`| Cadena &#124; Número &#124; Booleano &#124; Objeto &#124; Matriz|El valor que se debe almacenar.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](../tutorial-api-requirement-sets.md)| 1.0|
|[Nivel de permisos mínimo](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restringido|
|Modo de Outlook aplicable| Redacción o lectura|