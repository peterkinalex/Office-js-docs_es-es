

# Método ProjectDocument.removeHandlerAsync
Quita de forma asíncrona un controlador de eventos para el evento cambiado de selección de tareas que hay en un objeto [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md).

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selección|
|**Agregado en**|1,0|

```js
Office.context.document.removeHandlerAsync(eventType[, options][, callback]);
```


## Parámetros
|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
|_eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|El tipo de evento que se debe quitar, como constante [EventType](../../reference/shared/eventtype-enumeration.md) o su valor de texto correspondiente. Necesario.<br/><br/>En la siguiente tabla se muestran los argumentos eventType válidos para un objeto [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md).<br/><br/><table><tr><th>Enumeración</th><th>Valor de texto</th></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp179836.aspx">Office.EventType.ResourceSelectionChanged</a></td><td>resourceSelectionChanged</td></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp179816.aspx">Office.EventType.TaskSelectionChanged</a></td><td>taskSelectionChanged</td></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp179839.aspx">Office.EventType.ViewSelectionChanged</a></td><td>viewSelectionChanged</td></tr></table>||
|_options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
|_asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
|_callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||


## Valor de devolución de llamada

Cuando la función _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el parámetro de la función de devolución de llamada.

Para el método **removeHandlerAsync**, el objeto devuelto [AsyncResult](../../reference/shared/asyncresult.md) contiene las siguientes propiedades.


|**Nombre**|**Descripción**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Datos pasados en el parámetro opcional _asyncContext_, si se usó el parámetro.|
|[error](../../reference/shared/asyncresult.error.md)|Información sobre el error si la propiedad **status** es igual a **failed**.|
|[estado](../../reference/shared/asyncresult.status.md)|Estado **succeeded** o **failed** de la llamada asincrónica.|
|[value](../../reference/shared/asyncresult.value.md)|**removeHandlerAsync** siempre devuelve **undefined**.|

## Ejemplo

El siguiente código de ejemplo utiliza [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) para agregar un controlador de eventos para el evento [ResourceSelectionChanged](../../reference/shared/projectdocument.resourceselectionchanged.event.md) y **removeHandlerAsync** para quitar el controlador.

Cuando se selecciona un recurso en una vista de recursos, el controlador muestra el GUID del recurso. Cuando se quita el controlador, no se muestra el GUID.

En el ejemplo se presupone que el complemento tiene una referencia a la biblioteca de jQuery y que el siguiente control de página se define en div de contenido en el cuerpo de la página.




```HTML
<input id="remove-handler" type="button" value="Remove handler" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            Office.context.document.addHandlerAsync(
                Office.EventType.ResourceSelectionChanged,
                getResourceGuid);
            $('#remove-handler').click(removeEventHandler);
        });
    };

    // Remove the event handler.
    function removeEventHandler() {
        Office.context.document.removeHandlerAsync(
            Office.EventType.ResourceSelectionChanged,
            {handler:getResourceGuid,
            asyncContext:'The handler is removed.'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#remove-handler').attr('disabled', 'disabled');
                    $('#message').html(result.asyncContext);
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


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|
|:-----|:-----|:-----|
|**Project**|v||

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|Selección|
|**Nivel de permisos mínimo**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad

|**Versión**|**Cambios**|
|:-----|:-----|
|1,0|Agregado|

## Vea también



#### Otros recursos


[Método addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)
[Enumeración EventType](../../reference/shared/eventtype-enumeration.md)
[Objeto ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)

