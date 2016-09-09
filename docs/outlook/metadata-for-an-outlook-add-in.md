
# Obtener y establecer los metadatos de complemento de un complemento de Outlook

Puede administrar datos personalizados del complemento de Outlook si usa uno de los métodos siguientes:

- Configuración de itinerancia, que administra datos personalizados en el buzón de un usuario.
    
- Propiedades personalizadas, que administran datos personalizados para un elemento en el buzón de un usuario.
    
Ambos dan acceso a los datos personalizados a los que solo se puede acceder mediante el complemento de Outlook, pero cada método almacena los datos por separado de los demás. Es decir, las propiedades personalizadas no pueden acceder a los datos almacenados a través de la configuración de movilidad y viceversa. Los datos se almacenan en el servidor de ese buzón de correo y son accesibles en sesiones de Outlook subsiguientes en todos los factores de forma que el complemento admite. 

## Datos personalizados por buzón: opciones de movilidad


Puede especificar los datos específicos del buzón de Exchange de un usuario mediante el objeto [RoamingSettings](../../reference/outlook/RoamingSettings.md). Los datos personales y las preferencias del usuario son ejemplos de estos datos. El complemento de correo puede obtener acceso a la configuración de movilidad cuando se mueve en cualquier dispositivo en el que pueda ejecutarse por diseño (escritorio, tableta o smartphone).

 Los cambios en estos datos se almacenan en una copia en memoria de los parámetros de la sesión actual de Outlook. Debe guardar la configuración de movilidad explícitamente tras su actualización para que esté disponible la próxima vez que el usuario abra el complemento, en el mismo dispositivo o cualquier otro compatible.


### Formato de las configuraciones de movilidad


Los datos de un objeto  **RoamingSettings** se almacenan como una cadena serializada de notación de objetos JavaScript (JSON). A continuación se muestra un ejemplo de la estructura, asumiendo que hay tres configuraciones de movilidad definidas con los nombres `add-in_setting_name_0`,  `add-in_setting_name_1` y `add-in_setting_name_2`.


```js
{
  "add-in_setting_name_0":"add-in_setting_value_0",
  "add-in_setting_name_1":"add-in_setting_value_1",
  "add-in_setting_name_2":"add-in_setting_value_2"
}
```


### Carga de la configuración de movilidad


Un complemento de correo normalmente carga la configuración de movilidad del controlador de eventos [Office.initialize](../../reference/shared/office.initialize.md). En el siguiente ejemplo de código JavaScript se muestra cómo cargar la configuración de itinerancia existente y obtener los valores de dos configuraciones ("customerName" y "customerBalance"):


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### Creación o asignación de una configuración de movilidad


Siguiendo con el ejemplo anterior, la siguiente función de JavaScript ( `setAddInSetting`) muestra cómo usar el método [RoamingSettings.set](../../reference/outlook/RoamingSettings.md) para establecer una opción denominada `cookie` con la fecha de hoy y conservar los datos mediante el método [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md) para volver a guardar todas las opciones de configuración de movilidad en el servidor. El método **set** crea la opción si no existe aún y asigna la configuración al valor especificado. El método **saveAsync** guarda la configuración de movilidad de manera asincrónica. Este ejemplo de código pasa un método de devolución de llamada ( `saveMyAddInSettingsCallback`) a  **saveAsync**. Cuando termina la llamada asincrónica, se llama a  `saveMyAddInSettingsCallback` mediante un parámetro, _asyncResult_. Este parámetro es un objeto [AsyncResult](../../reference/outlook/simple-types.md) que contiene el resultado y los detalles de la llamada asincrónica. Puede usar el parámetro opcional _userContext_ para pasar información de estado desde la llamada asincrónica a la función de devolución de llamada.


```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### Supresión de la configuración de movilidad


Además, para ampliar los ejemplos anteriores, la siguiente función de JavaScript,  `removeAddInSetting`, muestra cómo usar el método [RoamingSettings.remove](../../reference/outlook/RoamingSettings.md) para quitar la opción `cookie` y volver a guardar todas las opciones de configuración de movilidad en el servidor de Exchange Server.


```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## Datos personalizados por cada elemento en un buzón de correo: propiedades personalizadas


Puede especificar datos concretos en un elemento en el buzón del usuario mediante el objeto [CustomProperties ](../../reference/outlook/CustomProperties.md). Por ejemplo, el complemento de correo podría clasificar determinados mensajes y anotar la categoría mediante una propiedad personalizada  `messageCategory`. O bien, si el complemento de correo crea citas a partir de sugerencias de reunión en un mensaje, se puede usar una propiedad personalizada para realizar un seguimiento de cada una de las citas. Esto garantiza que si el usuario vuelve a abrir el mensaje, el complemento de correo no proponga crear la cita una segunda vez.

Tal como ocurre con la configuración de movilidad, los cambios en las propiedades personalizadas se almacenan en las copias en la memoria de las propiedades para la sesión actual de Outlook. Para asegurarse de que estas propiedades personalizadas estarán disponibles en la siguiente sesión, guarde todas las propiedades personalizadas en el servidor.

Solo se puede tener acceso a estas propiedades personalizadas específicas del elemento y del complemento mediante el objeto  **CustomProperties**. Estas propiedades son diferentes de las propiedades personalizadas basadas en MAPI, [UserProperties ](http://msdn.microsoft.com/library/20b49c86-d74f-9bda-382c-559af278c148%28Office.15%29.aspx), en el modelo de objetos de Outlook y las propiedades extendidas en Servicios Web Exchange (EWS). No se puede tener acceso a  **CustomProperties** mediante el modelo de objetos de Outlook o EWS.

Sin embargo, un complemento de correo puede obtener las propiedades extendidas basadas en MAPI mediante la operación de EWS [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx). Obtenga acceso a  **GetItem** en el servidor con un token de devolución de llamada o en el cliente mediante el método [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md). En la solicitud  **GetItem**, especifique las propiedades extendidas personalizadas que necesita en un conjunto de propiedades. Un complemento de correo también puede usar  **makeEwsRequestAsync** y las operaciones de EWS [CreateItem](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx) y [UpdateItem](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) para crear y modificar propiedades extendidas.




### Uso de propiedades personalizadas


Para poder usar propiedades personalizadas, debe cargarlas llamando al método [loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md). Si ya hay propiedades personalizadas establecidas para el elemento actual, se cargan desde el servidor de Exchange en este momento. Después de crear el contenedor de propiedades, podrá usar los métodos [set](../../reference/outlook/CustomProperties.md) y [get](../../reference/outlook/CustomProperties.md) para agregar y recuperar propiedades personalizadas. Para guardar los cambios que haga en el contenedor de propiedades y conservarlos en el servidor de Exchange, use el método [saveAsync](../../reference/outlook/CustomProperties.md).


 >**Nota**  Dado que Outlook para Mac no almacena las propiedades personalizadas en caché, si la red del usuario presentara algún error, los complementos de correo en Outlook para Mac no podrían obtener acceso a sus propiedades personalizadas.


### Ejemplo de propiedades personalizadas


En el siguiente ejemplo se muestra un conjunto simplificado de métodos para un complemento de Outlook que usa propiedades personalizadas. Puede usar este ejemplo como punto de partida de su complemento que usa propiedades personalizadas. 

Este ejemplo incluye los métodos siguientes:


- [Office.initialize](../../reference/shared/office.initialize.md): inicializa el complemento y carga el contenedor de propiedades personalizadas desde el servidor Exchange.
    
-  **customPropsCallback**: obtiene el contenedor de propiedades personalizadas que devuelve el servidor y lo guarda para usarlo posteriormente.
    
-  **updateProperty**: establece o actualiza una propiedad concreta y, después, guarda los cambios en el servidor.
    
-  **removeProperty**: quita una propiedad concreta del contenedor de propiedades y, después, guarda la eliminación en el servidor.
    



```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


## Recursos adicionales

    
- [Información general sobre MAPI (propiedad)](http://msdn.microsoft.com/library/02e5b23f-1bdb-4fbf-a27d-e3301a359573%28Office.15%29.aspx)
    
- [Resumen de las propiedades de Outlook](http://msdn.microsoft.com/library/242c9e89-a0c5-ff89-0d2a-410bd42a3461%28Office.15%29.aspx)
    
- [Llamar a servicios web desde un complemento de Outlook](../outlook/web-services.md)
    
- [Propiedades y propiedades extendidas de EWS en Exchange](http://msdn.microsoft.com/library/68623048-060e-4602-b3fa-62617a94cf72%28Office.15%29.aspx)
    
- [Conjuntos de propiedades y respuesta de formas de EWS en Exchange](http://msdn.microsoft.com/library/04a29804-6067-48e7-9f5c-534e253a230e%28Office.15%29.aspx)
    


