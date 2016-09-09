
# Conservación de la configuración y del estado de los complementos

Las Complementos de Office son, esencialmente, aplicaciones web que se ejecutan en el entorno sin estado de un control del explorador. Como resultado, es posible que su aplicación requiera conservar datos para mantener la continuidad de determinadas operaciones o características entre sesiones de uso. Por ejemplo, es posible que su complemento tenga configuraciones personalizadas u otros valores que necesite guardar y recargar la próxima vez que se inicialice, como puede ser la vista preferida de un usuario o una ubicación predeterminada.

Para ello, puede:


- Usar miembros de la API de JavaScript para Office que almacenan datos como pares nombre-valor en un contenedor de propiedades almacenado en una ubicación que depende del tipo de complemento.
    
- Usar técnicas proporcionadas por el control del explorador subyacente: cookies de explorador o almacenamiento web HTML5 ([localStorage](http://msdn.microsoft.com/en-us/library/cc848902%28v=vs.85%29.aspx) o [sessionStorage](http://msdn.microsoft.com/en-us/library/cc197020%28v=vs.85%29.aspx)).
    
Este artículo se centra en cómo usar la API de JavaScript para Office para conservar el estado de los complementos. Si quiere ejemplos del uso de cookies del explorador y de almacenamiento web, vea [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## Conservación de la configuración y el estado de los complementos con la API de JavaScript para Office


La API de JavaScript para Office proporciona los objetos [Settings](../../reference/shared/settings.md), [RoamingSettings](../../reference/outlook/RoamingSettings.md) y [CustomProperties](../../reference/outlook/CustomProperties.md) para guardar estados de complementos en distintas sesiones, como se describe en la siguiente tabla. En todos los casos, los valores de configuración guardados se asocian con el [identificador](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx) del complemento que los creó.



|**Object**|**Tipo de complemento compatible**|**Ubicación del almacenamiento**|**Hosts de Office compatibles**|
|:-----|:-----|:-----|:-----|
|[Configuración](../../reference/shared/settings.md)|panel de tareas y de contenido|El documento, la hoja de cálculo o la presentación con los que se use el complemento. Las opciones de configuración de los complementos de panel de tareas y de contenido están disponibles para el complemento que los creó en el documento donde se guardaron. **Importante:** No almacene contraseñas ni otra información de identificación personal (DCP) confidencial con el objeto **Settings**. Los datos guardados no son visibles para los usuarios finales, pero se almacenan como parte del documento, que sí es accesible si se lee directamente el formato de archivo del documento. Es necesario que limite el uso de DCP por el complemento y que almacene la DCP que necesite el complemento solo en el servidor donde se hospeda el complemento como un recurso protegido frente a los usuarios.|Word, Excel o PowerPoint **Nota:** Los complementos de panel de tareas para Project 2013 no son compatibles con la API de **Settings** para almacenar el estado o la configuración de los complementos. Pero, para los complementos que se ejecutan en Project (y en otras aplicaciones host de Office), se pueden usar técnicas como las cookies del explorador o el almacenamiento web. Para más información sobre estas técnicas, vea [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](../../reference/outlook/RoamingSettings.md)|Outlook|El buzón del usuario en el servidor Exchange en el que está instalado el complemento.Dado que esta configuración se almacena en el buzón de servidor del usuario, podrá "moverse" junto con el usuario y estará a disposición del complemento cuando se use en el contexto de cualquier aplicación host cliente o explorador que tenga acceso al buzón de correo del usuario. La configuración de movilidad del complemento de Outlook solo está disponible para la aplicación que la creó y únicamente desde el buzón en el que está instalado el complemento.|Outlook|
|[CustomProperties](../../reference/outlook/CustomProperties.md)|Outlook|El elemento de mensaje, cita o convocatoria de reunión con que está trabajando el complemento. Las propiedades personalizadas del elemento del complemento de Outlook solo están disponibles para el complemento que las creó y únicamente desde el elemento en el que están guardadas.|Outlook|

## Los datos de la configuración se administran en la memoria en tiempo de ejecución


Internamente, los datos del contenedor de propiedades a los que se tiene acceso con los objetos  **Settings**,  **CustomProperties** o **RoamingSettings** se almacenan como un objeto de notación JavaScript (JSON) serializado que contiene pares nombre-valor. El nombre (clave) de cada valor debe ser una **string** y el valor almacenado puede ser una **string**, un  **number**, una  **date** o un **object** de JavaScript, pero no una **function**.

Este ejemplo de la estructura del contenedor de propiedades contiene tres valores de  **string** definidos: `firstName`,  `location` y `defaultView`.




```
{
"firstName":"Erik",
"location":"98052",
"defaultView":"basic"
}
```

Después de que se haya guardado el contenedor de propiedades de configuración durante la sesión anterior del complemento, puede cargarse al iniciar la aplicación o en cualquier momento posterior, durante la sesión actual del complemento. Durante la sesión, la configuración se administra por completo en la memoria mediante los métodos  **get**,  **set** y **remove** del objeto que corresponda al tipo de configuración que está creando ( **Settings**,  **CustomProperties** o **RoamingSettings**). 


 >**Importante**  Para conservar las adiciones, actualizaciones o eliminaciones realizadas durante la sesión actual del complemento en la ubicación de almacenamiento, debe llamar al método  **saveAsync** del objeto correspondiente que se use para trabajar con ese tipo de configuración. Los métodos **get**,  **set** y **remove** solamente operan sobre la copia del contenedor de propiedades de configuración que se encuentra en la memoria. Si su complemento se cierra sin llamar a **saveAsync**, se perderán todos los cambios realizados en la configuración durante esa sesión. 


## Procedimiento para guardar la configuración y el estado de los complementos por cada documento de los complementos de panel de tareas y de contenido


Para conservar la configuración personalizada o el estado de un complemento de panel de tareas o de contenido para Word, Excel o PowerPoint, es necesario usar el objeto [Settings](../../reference/shared/settings.md) y sus métodos. El contenedor de propiedades que se crea con los métodos del objeto **Settings** solo está disponible para la instancia del complemento de panel de tareas o de contenido que lo creó y solo desde el documento donde se guarde.

El objeto  **Settings** se carga automáticamente como parte del objeto [Document](../../reference/shared/document.md), y está disponible cuando se activa el complemento de contenido o panel de tareas. Una vez que se formuló la instancia del objeto  **Document**, puede acceder al objeto  **Settings** con la propiedad [settings](../../reference/shared/document.settings.md) del objeto **Document**. Durante la sesión, puede simplemente usar los métodos  **Settings.get**,  **Settings.set** y **Settings.remove** para leer, escribir y quitar la configuración y el estado de la aplicación de la copia en la memoria del contenedor de propiedades.

Como los métodos Set y Remove solo realizan operaciones en la copia en memoria del contenedor de propiedades de configuración, para volver a guardar la configuración nueva o modificada en el documento con el que está asociado el complemento es necesario llamar al método [Settings.saveAsync](../../reference/shared/settings.saveasync.md).


### Creación o actualización de un valor de configuración

En el ejemplo de código siguiente se muestra cómo usar el método [Settings.set](../../reference/shared/settings.set.md) para crear una configuración llamada `'themeColor'` con el valor `'green'`. El primer parámetro del método Set es el nombre (_name_) o id., que distingue mayúsculas de minúsculas, de la configuración que se va a establecer o crear. El segundo parámetro es el valor (_value_) de la configuración.


```
Office.context.document.settings.set('themeColor', 'green');
```

 Si la configuración no existe aún, se crea una con el nombre especificado o, si ya existe, se actualiza su valor. Use el método **Settings.saveAsync** para conservar la configuración nueva o actualizada en el documento.


### Obtención del valor de una configuración

En el ejemplo siguiente se muestra cómo usar el método [Settings.get](../../reference/shared/settings.get.md) para obtener el valor de una configuración llamada "themeColor". El único parámetro del método **get** es el nombre (_name_) de la configuración, que distingue mayúsculas de minúsculas.


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 El método **get** devuelve el valor que se guardó anteriormente para la configuración de _name_ que se pasó. Si la configuración no existe, el método devuelve **null**.


### Quitar una configuración

En el ejemplo siguiente se muestra cómo usar el método [Settings.remove](../../reference/shared/settings.removehandlerasync.md) para quitar una configuración llamada "themeColor". El único parámetro del método **remove** es el nombre (_name_) de la configuración, que distingue mayúsculas de minúsculas.


```
Office.context.document.settings.remove('themeColor');
```

No se producirá ninguna acción si la configuración no existe. Use el método  **Settings.saveAsync** para trasladar la supresión de la configuración al documento.


### Almacenamiento de la configuración

Para guardar las adiciones, cambios o eliminaciones que realice el complemento en la copia en memoria del contenedor de propiedades de la configuración durante la sesión actual, es necesario llamar al método [Settings.saveAsync](../../reference/shared/settings.saveasync.md) para almacenarlos en el documento. El único parámetro del método **saveAsync** es _callback_, que es una función de devolución de llamada con un único parámetro. 


```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

La función anónima que se pasa al método  **saveAsync** como parámetro de _callback_ se ejecuta cuando se completa la operación. El parámetro _asyncResult_ de la devolución de llamada proporciona acceso a un objeto **AsyncResult** que contiene el estado de la operación. En el ejemplo, la función comprueba la propiedad **AsyncResult.status** para ver si la operación de almacenaje se realizó correctamente o no, para a continuación mostrar el resultado en la página del complemento.


## Procedimiento para guardar la configuración en el buzón del usuario para complementos de Outlook como valores de configuración del servicio de movilidad


Un complemento de Outlook puede usar el objeto [RoamingSettings](../../reference/outlook/RoamingSettings.md) para guardar los datos de estado y configuración de la aplicación que son específicos del buzón del usuario. A estos datos solo tiene acceso el complemento de Outlook en nombre del usuario que ejecuta la aplicación. Los datos se almacenan en el buzón del usuario del servidor Exchange y el usuario podrá obtener acceso a ellos cuando inicie sesión en la cuenta y ejecute el complemento de Outlook.


### Carga de la configuración de movilidad


Normalmente, una aplicación de Outlook carga las configuraciones de movilidad en el controlador de eventos [Office.initialize](../../reference/shared/office.initialize.md). El siguiente ejemplo de código JavaScript muestra cómo cargar configuraciones de movilidad existentes.


```
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}

```


### Creación o asignación de una configuración de movilidad


Si continuamos con el ejemplo anterior, la siguiente función  `setAppSetting` muestra cómo usar el método [RoamingSettings.set](../../reference/outlook/RoamingSettings.md) para crear o actualizar una configuración con el nombre `cookie` y la fecha de hoy. Luego guarda todos los valores de configuración de movilidad en el servidor Exchange con el método [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md).


```
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

El método  **saveAsync** guarda las configuraciones de movilidad de forma asincrónica y toma una función de devolución de llamada opcional. Este ejemplo de código pasa un método de devolución de llamada denominado `saveMyAppSettingsCallback` al método **saveAsync**. Cuando la llamada asincrónica regresa, el parámetro  _asyncResult_ de la función `saveMyAppSettingsCallback` proporciona acceso a un objeto [AsyncResult](../../reference/outlook/simple-types.md) que se puede usar para determinar el éxito o el fracaso de la operación mediante la propiedad **AsyncResult.status**.


### Supresión de la configuración de movilidad


Como ampliación de los ejemplos anteriores, la siguiente función  `removeAppSetting` muestra cómo usar el método [RoamingSettings.remove](../../reference/outlook/RoamingSettings.md) para quitar la configuración `cookie` y guardar todas las configuraciones de movilidad en Exchange Server.


```
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## Procedimiento para guardar la configuración por cada elemento de los complementos de Outlook como propiedades personalizadas


Las propiedades personalizadas permiten a su complemento de Outlook almacenar información sobre un elemento con el que está trabajando. Por ejemplo, si el complemento de Outlook crea una cita para una sugerencia de reunión en un mensaje, puede usar las propiedades personalizadas para almacenar el hecho de que se creó la reunión. Esto le asegura que, si se vuelve a abrir el mensaje, el complemento de Outlook no volverá a ofrecer crear otra cita.

Si quiere usar propiedades personalizadas para un elemento de convocatoria de reunión, una cita o un mensaje determinados, debe cargar las propiedades en memoria llamando al método [loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md) del objeto **Item**. Si alguna de las propiedades personalizadas ya está establecida para el elemento actual, estas se cargarán desde el servidor Exchange en este punto. Una vez cargadas las propiedades, use los métodos [set](../../reference/outlook/CustomProperties.md) y [get](../../reference/outlook/RoamingSettings.md) del objeto **CustomProperties** para agregar, actualizar y recuperar propiedades en la memoria. Para guardar los cambios realizados en las propiedades personalizadas del elemento, debe usar el método [saveAsync](../../reference/outlook/CustomProperties.md) para conservar los cambios del elemento en el servidor Exchange.


### Ejemplo de propiedades personalizadas

En el siguiente ejemplo se muestra un conjunto simplificado de funciones para un complemento de Outlook que usa propiedades personalizadas. Puede usar este ejemplo como punto de partida de su complemento de Outlook que usa propiedades personalizadas. 

Los complementos de Outlook que usan estas funciones recuperan las propiedades personalizadas al llamar al método  **get** en la variable `_customProps`, como se muestra en el ejemplo siguiente.




```
var property = _customProps.get("propertyName");
```

Este ejemplo incluye las funciones siguientes:



|**Nombre de la función**|**Descripción**|
|:-----|:-----|
| `Office.initialize`|Inicializa el complemento y carga las propiedades personalizadas para el elemento actual desde el servidor Exchange.|
| `customPropsCallback`|Obtiene las propiedades personalizadas del servidor Exchange y las guarda para usarlas más adelante.|
| `updateProperty`|Establece o actualiza una propiedad concreta y luego guarda los cambios realizados en el servidor Exchange.|
| `removeProperty`|Elimina una propiedad concreta y luego conserva la eliminación en el servidor Exchange.|
| `saveCallback`|Devolución de llamada para llamadas al método  **saveAsync** en las funciones `updateProperty` y `removeProperty`.|



```
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change 
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal 
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method. 
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```


## Recursos adicionales



- [Información sobre la API de JavaScript para Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Complementos de Outlook](../outlook/outlook-add-ins.md)
    
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
