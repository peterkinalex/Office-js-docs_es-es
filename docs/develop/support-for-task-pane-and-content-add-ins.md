
# La API de JavaScript para Office es compatible con complementos de contenido y panel de tareas en Office 2013


Puede usar la [API de JavaScript para Office](../../reference/javascript-api-for-office.md) para crear complementos de panel de tareas o de contenido para aplicaciones host de Office 2013. Los objetos y los métodos compatibles con los complementos de contenido y panel de tareas se dividen en las categorías siguientes:


1. **Objetos comunes compartidos con otros complementos de Office.** Estos objetos son [Office](../../reference/shared/office.md), [Context](../../reference/shared/office.context.md), y [AsyncResult](../../reference/shared/asyncresult.md). El objeto **Office** es el objeto raíz de la API de JavaScript para Office. El objeto **Context** representa el entorno de tiempo de ejecución del complemento. Los objetos **Office** y **Context** son fundamentales para cualquier complemento de Office. El objeto **AsyncResult** representa los resultados de una operación asincrónica, como los datos devueltos al método **getSelectedDataAsync**, que lee lo que ha seleccionado un usuario en un documento.
    
2.  **El objeto Document.** La mayoría de la API disponible para los complementos de contenido y panel de tareas se expone con los métodos, propiedades y eventos del objeto [Document](../../reference/shared/document.md). Un complemento de contenido o panel de tareas puede usar la propiedad [Office.context.document](../../reference/shared/office.context.document.md) para tener acceso al objeto **Document** y, a través de este, puede tener acceso a los miembros principales de la API para trabajar con datos y documentos, como los objetos [Bindings](../../reference/shared/bindings.bindings.md) y [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md), así como los métodos [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md), [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) y [getFileAsync](../../reference/shared/document.getfileasync.md). El objeto **Document** también proporciona la propiedad [mode](../../reference/shared/document.mode.md) (para determinar si un documento es de solo lectura o está en el modo de edición), la propiedad [url](../../reference/shared/document.url.md) (para obtener la URL del documento actual) y acceso al objeto [Settings](../../reference/shared/settings.md). El objeto **Document** también permite agregar controladores de eventos para el evento [SelectionChanged](../../reference/shared/document.selectionchanged.event.md), con el que se puede detectar cuando un usuario cambia su selección en el documento.
    
   Un complemento de contenido o panel de tareas solo puede tener acceso al objeto **Document** cuando se haya cargado el DOM y el entorno de tiempo de ejecución, normalmente en el controlador de eventos del evento [Office.initialize](../../reference/shared/office.initialize.md). Para obtener información sobre el flujo de eventos cuando se inicia un complemento y cómo comprobar si el DOM y el tiempo de ejecución se cargaron correctamente, vea [Cargar el DOM y el entorno de tiempo de ejecución](../../docs/develop/loading-the-dom-and-runtime-environment.md).
    
3.  **Objetos para trabajar con características específicas.** Para trabajar con características específicas de la API, use los siguientes objetos y métodos:
    
    - Los métodos del objeto [Bindings](../../reference/shared/bindings.bindings.md) para crear u obtener enlaces, y los métodos y propiedades del objeto [Binding](../../reference/shared/binding.md) para trabajar con datos.
    
    - [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md), [CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md) y los objetos asociados para crear y manipular fragmentos XML personalizados en documentos de Word.
    
    - Los objetos [File](../../reference/shared/file.md) y [Slice](../../reference/shared/slice.md) para crear una copia de todo el documento, dividirlo en fragmentos o "segmentos" y, después, leer o transmitir datos en esos segmentos.
    
    - El objeto [Settings](../../reference/shared/settings.md) para guardar datos personalizados, como preferencias de usuario y el estado del complemento.
    

 >**Importante** Algunos miembros de la API no son compatibles en todas las aplicaciones de Office que pueden hospedar complementos de contenido y panel de tareas. Para determinar qué miembros son compatibles, vea uno de los vínculos siguientes:

Para consultar un resumen de la API de JavaScript para Office en diferentes aplicaciones host de Office, vea [Información sobre la API de JavaScript para Office](../../docs/develop/understanding-the-javascript-api-for-office.md).


## Lectura y escritura en una selección activa

Puede leer o escribir en la selección actual de un usuario en un documento, una hoja de cálculo o una presentación. En función de la aplicación host del complemento, puede especificar el tipo de estructura de datos donde quiere leer o escribir como un parámetro de los métodos [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) y [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) del objeto [Document](../../reference/shared/document.md). Por ejemplo, puede especificar cualquier tipo de datos para Word (texto, HTML, datos tabulares u Office Open XML), texto y datos tabulares para Excel y texto para PowerPoint y Project. También puede crear controladores de eventos para detectar cambios en la selección del usuario. El ejemplo siguiente obtiene datos de una selección como texto con el método **getSelectedDataAsync**.


```js
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        }
        else {
            write('Selected data: ' + asyncResult.value);
        }
    });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

Para obtener más información y ejemplos, vea [Leer y escribir datos en la selección activa de un documento u hoja de cálculo](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## Enlaces a una región de un documento u hoja de cálculo

Puede usar los métodos **getSelectedDataAsync** y **setSelectedDataAsync** para leer o escribir en la selección *actual* del usuario en un documento, hoja de cálculo o presentación. Pero, si quiere tener acceso a la misma región en un documento en varias sesiones de ejecución del complemento sin que sea necesario que el usuario realice una selección, primero tiene que enlazar esa región. También se puede suscribir a eventos de cambios de datos y selección para esa región enlazada.

Los enlaces pueden agregarse con los métodos [addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md), [addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md) y [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) del objeto [Bindings](../../reference/shared/bindings.bindings.md). Estos métodos devuelven un identificador que permite acceder a los datos del enlace o suscribirse a los eventos de cambio de selección o datos correspondientes.

Este es un ejemplo que agrega un enlace al texto seleccionado actualmente en un documento con el método **Bindings.addFromSelectionAsync**.



```js
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Para obtener más información y ejemplos, vea [Enlazar a regiones en un documento u hoja de cálculo](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


## Obtención de documentos completos

Si el complemento de panel de tareas se ejecuta en PowerPoint o Word, puede usar los métodos [Document.getFileAsync](../../reference/shared/document.getfileasync.md), [File.getSliceAsync](../../reference/shared/file.getsliceasync.md) y [File.closeAsync](../../reference/shared/file.closeasync.md) para obtener una presentación o documento completo.

Al realizar una llamada a **Document.getFileAsync**, se obtiene una copia del documento en el objeto [File](../../reference/shared/file.md). El objeto **File** proporciona acceso al documento en "fragmentos" representados como objetos [Slice](../../reference/shared/document.md). Al realizar una llamada a **getFileAsync**, se puede especificar el tipo de archivo (texto o formato Office Open XML comprimido) y el tamaño de los segmentos (hasta 4 MB). Después, para tener acceso al contenido del objeto **File**, se realiza una llamada a **File.getSliceAsync**, que devuelve los datos sin procesar en la propiedad [Slice.data](../../reference/shared/slice.data.md). Si especificó el formato comprimido, recibirá los datos del archivo como una matriz de bytes. Si transmite el archivo a un servicio web, puede transformar los datos sin formato comprimidos en una cadena con codificación base64 antes del envío. Por último, cuando termine de obtener segmentos del archivo, use el método **File.closeAsync** para cerrar el documento.

Para más información, vea cómo [obtener el documento completo de un complemento para PowerPoint o Word](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md). 


## Leer y escribir en fragmentos XML personalizados de un documento de Word

Si usa el formato de archivo Office Open XML y controles de contenido, puede agregar fragmentos XML personalizados a un documento de Word y enlazar elementos de los fragmentos XML a controles de contenido en el documento. Al abrir el documento, Word lee y rellena automáticamente los controles de contenido enlazados con datos de los fragmentos XML personalizados. Los usuarios también pueden escribir datos en los controles de contenido y, cuando el usuario guarde el documento, los datos en los controles se guardarán en los fragmentos XML enlazados. Los complementos de panel de tareas para Word pueden usar la propiedad [Document.customXmlParts](../../reference/shared/document.customxmlparts.md) y los objetos [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md), [CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md) y [CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md) para leer y escribir datos de forma dinámica en el documento.

Los fragmentos XML personalizados se pueden asociar con espacios de nombres. Para obtener datos de los fragmentos XML personalizados en un espacio de nombres, use el método [CustomXmlParts.getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md).

También puede usar el método [CustomXmlParts.getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md) para tener acceso a los fragmentos XML personalizados por sus GUID correspondientes. Después de obtener un fragmento XML personalizado, use el método [CustomXmlPart.getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md) para obtener los datos XML.

Para agregar un nuevo fragmento XML personalizado a un documento, use la propiedad **Document.customXmlParts** para obtener los fragmentos XML personalizados que están en el documento y realice una llamada al método [CustomXmlParts.addAsync](../../reference/shared/customxmlparts.addasync.md).

Para más información sobre cómo trabajar con fragmentos XML personalizados desde un complemento de panel de tareas, vea [Crear mejores complementos para Word con Office Open XML](../../docs/word/create-better-add-ins-for-word-with-office-open-xml.md).


## Conservación de la configuración de complementos


A menudo es preciso guardar datos personalizados para el complemento, como las preferencias del usuario o el estado del complemento, y tener acceso a esos datos la siguiente vez que se abre el complemento. Puede usar técnicas comunes de programación web para guardar estos datos, como cookies del navegador o almacenamiento web HTML 5. O bien, si el complemento se ejecuta en Excel, PowerPoint o Word, puede usar los métodos del objeto [Settings](../../reference/shared/settings.md). Los datos creados con el objeto **Settings** se almacenan en la hoja de cálculo, presentación o documento en el que se ha insertado y con el que se ha guardado el complemento. Estos datos están disponibles solo para el complemento que los ha creado.

Para evitar ciclos de ida y vuelta al servidor donde se almacena el documento, los datos creados con el objeto **Settings** se administran en la memoria en el tiempo de ejecución. Los datos de configuración guardados anteriormente se cargan en la memoria cuando se inicializa el complemento y los cambios en esos datos solo se vuelven a guardar en el documento cuando realiza una llamada al método [Settings.saveAsync](../../reference/shared/settings.saveasync.md). Internamente, los datos se almacenan en un objeto JSON en serie como pares de nombres/valores. Use los métodos [get](../../reference/shared/settings.get.md), [set](../../reference/shared/settings.set.md) y [remove](../../reference/shared/settings.removehandlerasync.md) del objeto **Settings** para leer, escribir y eliminar elementos de la copia almacenada en la memoria de los datos. En la siguiente línea de código se muestra cómo crear una configuración llamada `themeColor` y establecer su valor en "green".




```js
Office.context.document.settings.set('themeColor', 'green');
```

Como los datos de configuración creados o eliminados con los métodos **set** y **remove** actúan en una copia en la memoria de los datos, es necesario realizar una llamada a **saveAsync** para guardar los cambios en los datos de configuración en el documento con el que trabaja el complemento.

Para más información sobre cómo trabajar con datos personalizados con los métodos del objeto **Settings**, vea [Guardar el estado y la configuración de un complemento](../../docs/develop/persisting-add-in-state-and-settings.md).


## Lectura de propiedades de un documento de proyecto

Si el complemento de panel de tareas se ejecuta en Project, podrá leer los datos de algunos campos de proyecto, recursos y campos de tarea del proyecto activo. Para ello, use los métodos y eventos del objeto [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md), que extiende el objeto **Document** para proporcionar una función adicional específica de Project.

Vea ejemplos de lectura de datos de Project en [Crear su primer complemento de panel de tareas para Project 2013 con un editor de texto](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## Modelo de permisos y administración

El complemento usa el elemento **Permissions** del manifiesto para solicitar permiso para tener acceso al nivel de la función que necesita de la API de JavaScript para Office. Por ejemplo, si el manifiesto necesita acceso de lectura/escritura al documento, en el manifiesto será necesario especificar `ReadWriteDocument` como el valor de texto en el elemento **Permissions**. Como los permisos existen para proteger la privacidad y la seguridad de los usuarios, se recomienda solicitar el nivel mínimo de permisos necesario para sus características. En el ejemplo siguiente se muestra cómo solicitar el permiso **ReadDocument** en el manifiesto del panel de tareas.


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

Para más información, vea [Solicitar permisos para usar la API en complementos de contenido y panel de tareas](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).


## Recursos adicionales


- [API de JavaScript para Office](../../reference/javascript-api-for-office.md)
    
- [Referencia de esquema para manifiestos de complementos de Office](http://msdn.microsoft.com/en-us/library/7e0cadc3-f613-8eb9-57ef-9032cbb97f92.aspx)
    
- [Solucionar errores de usuario con los complementos de Office](../../docs/testing/testing-and-troubleshooting.md)
    
