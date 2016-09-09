
# Método ProjectDocument.addHandlerAsync
Agrega de forma asíncrona un controlador de evento para un evento de cambio en un objeto [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md).

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selección|
|**Agregado en**|1,0|

```
Office.context.document.addHandlerAsync(eventType, handler[, options][, callback]);
```


## Parámetros



|**Nombre**|**Tipo**|**Descripción**|
|:-----|:-----|:-----|
| _eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|El tipo de evento que se debe agregar, como constante [EventType](../../reference/shared/eventtype-enumeration.md) o su valor de texto correspondiente. Necesario. En la siguiente tabla se muestran los argumentos _eventType_ válidos para un objeto [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md).<table><tr><td>**Enumeración**</td><td>**Valor de texto**</td></tr><tr><td>[Office.EventType.ResourceSelectionChanged](../../reference/shared/projectdocument.resourceselectionchanged.event.md)</td><td>resourceSelectionChanged</td></tr><tr><td>[Office.EventType.TaskSelectionChanged](../../reference/shared/projectdocument.taskselectionchanged.event.md)</td><td>taskSelectionChanged</td></tr><tr><td>[Office.EventType.ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md)</td><td>viewSelectionChanged</td></tr></table>|
| _handler_|**función**|El nombre del controlador de eventos. Obligatorio.|
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):|
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.|
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.|

## Valor de devolución de llamada

Cuando la función _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el parámetro de la función de devolución de llamada.

Para el método **addHandlerAsync**, el objeto devuelto [AsyncResult](../../reference/shared/asyncresult.md) contiene las siguientes propiedades:


****


|**Nombre**|**Descripción**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Datos pasados en el parámetro opcional _asyncContext_, si se usó el parámetro.|
|[error](../../reference/shared/asyncresult.error.md)|Información sobre el error si la propiedad **status** es igual a **failed**.|
|[estado](../../reference/shared/asyncresult.status.md)|Estado **succeeded** o **failed** de la llamada asincrónica.|
|[value](../../reference/shared/asyncresult.value.md)|**addHandlerAsync** siempre devuelve **undefined**.|

## Ejemplo

En el siguiente código de ejemplo se usa **addHandlerAsync** para agregar un controlador de eventos para el evento [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md).

Cuando se cambia la vista activa, el controlador comprueba el tipo de vista. Habilita un botón si la vista es una vista de recursos y deshabilita el botón si no es una vista de recursos. Obtiene el GUID del recurso seleccionado al seleccionar el botón y lo muestra en el complemento.

En el ejemplo se asume que el complemento tiene una referencia a la biblioteca de jQuery y que los controles de la siguiente página se definen en el div de contenido del cuerpo de la página.




```HTML
<input id="get-info" type="button" value="Get info" disabled="disabled" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            // Add a ViewSelectionChanged event handler.
            Office.context.document.addHandlerAsync(
                Office.EventType.ViewSelectionChanged,
                getActiveView);
            $('#get-info').click(getResourceGuid);

            // This example calls the handler on page load to get the active view
            // of the default page.
            getActiveView();
        });
    };

    // Activate the button based on the active view type of the document.
    // This is the ViewSelectionChanged event handler.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var viewType = result.value.viewType;
                    if (viewType == 6 ||   // ResourceForm
                        viewType == 7 ||   // ResourceSheet
                        viewType == 8 ||   // ResourceGraph
                        viewType == 15) {  // ResourceUsage
                        $('#get-info').removeAttr('disabled');
                    }
                    else {
                        $('#get-info').attr('disabled', 'disabled');
                    }
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, viewType);
                    $('#message').html(output);
                }
            }
        );
    }

    // Get the GUID of the currently selected resource and display it in the add-in.
    function getResourceGuid() {
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html('Resource GUID: ' + result.value);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```

Para obtener un ejemplo de código completo en el que se muestre cómo usar un controlador de eventos [TaskSelectionChanged](../../reference/shared/projectdocument.taskselectionchanged.event.md) en un complemento de Project, consulte [Create your first task pane app for Project by using a text editor](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md) (Crear su primer complemento de panel de tareas para Project con un editor de texto).


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|
|:-----|:-----|:-----|
|**Project**|v||

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **||
|**Nivel de permisos mínimo**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1,0|Agregado|

## Vea también



#### Otros recursos


[TaskSelectionChanged event](../../reference/shared/projectdocument.taskselectionchanged.event.md)

[Método removeHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)

[Objeto ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
