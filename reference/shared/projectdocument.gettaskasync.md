

# <a name="projectdocument.gettaskasync-method"></a>Método ProjectDocument.getTaskAsync
Obtiene de forma asincrónica el nombre de la tarea especificada, los recursos asignados y el identificador de la tarea en la lista sincronizada de tareas de SharePoint.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selección|
|**Agregado en**|1.0|

```js
Office.context.document.getTaskAsync(taskId [,options][, callback]);
```


## <a name="parameters"></a>Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _taskId_|**string**|El GUID de la tarea. Obligatorio.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el parámetro de la función de devolución de llamada.

Para el método **getTaskAsync**, el objeto devuelto [AsyncResult](../../reference/shared/asyncresult.md) contiene las siguientes propiedades.


****


|**Nombre**|**Descripción**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Datos pasados en el parámetro opcional _asyncContext_, si se usó el parámetro.|
|[error](../../reference/shared/asyncresult.error.md)|Información sobre el error si la propiedad **status** es igual a **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Estado **succeeded** o **failed** de la llamada asincrónica.|
|[value](../../reference/shared/asyncresult.value.md)|Contiene las siguientes propiedades:<br/><br/><ul><li><b>taskName</b>: el nombre de la tarea.</li><li><b>wssTaskId</b>: el id. de la tarea de la lista de tareas sincronizadas de SharePoint. Si el proyecto no está sincronizado con una lista de tareas de SharePoint, el valor es <b>0</b>.</li><li><b>resourceNames</b>: lista separada por comas de los nombres de los recursos asignados a la tarea.</li></ul>|

## <a name="remarks"></a>Observaciones

Antes de llamar al método **getTaskAsync**, llame al método [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) para obtener el GUID de la tarea.


## <a name="example"></a>Ejemplo

El siguiente código de ejemplo llama a [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) para obtener el GUID de la tarea que está seleccionada. A continuación, llama a **getTaskAsync** para obtener las propiedades para la tarea que están disponibles desde la API de JavaScript para Office.

En el ejemplo se asume que el complemento tiene una referencia a la biblioteca de jQuery y que los controles de la siguiente página se definen en el div de contenido del cuerpo de la página.




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            $('#get-info').click(getTaskInfo);
        });
    };

    // Get the GUID of the task, and then get local task properties.
    function getTaskInfo() {
        getTaskGuid().then(
            function (data) {
                getTaskProperties(data);
            }
        );
    }

    // Get the GUID of the selected task.
    function getTaskGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedTaskAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Get local properties for the selected task, and then display it in the add-in.
    function getTaskProperties(taskGuid) {
        Office.context.document.getTaskAsync(
            taskGuid,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var taskInfo = result.value;
                    var output = String.format(
                        'Name: {0}<br/>GUID: {1}<br/>SharePoint task ID: {2}<br/>Resource names: {3}',
                        taskInfo.taskName, taskGuid, taskInfo.wssTaskId, taskInfo.resourceNames);
                    $('#message').html(output);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();

```


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|
|:-----|:-----|:-----|
|**Project**|v||

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|Selección|
|**Nivel de permisos mínimo**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



|**Versión**|**Cambios**|
|:-----|:-----|
|1.0|Agregado|

## <a name="see-also"></a>Vea también



#### <a name="other-resources"></a>Otros recursos


[Método getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)
[Objeto AsyncResult](../../reference/shared/asyncresult.md)
[Objeto ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
