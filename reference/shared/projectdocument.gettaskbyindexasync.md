

# Método ProjectDocument.getTaskByIndexAsync
Obtiene de forma asincrónica el GUID de la tarea que tiene el índice especificado en la colección de tareas.

**Importante:** Esta API solo funciona en Project 2016 para el escritorio de Windows.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selección|
|**Agregado en**|1.1|

```js
Office.context.document.getTaskByIndexAsync(taskIndex[, options][, callback]);
```


## Parámetros

_taskIndex_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **número**

&nbsp;&nbsp;&nbsp;&nbsp;Índice de la tarea en la colección de tareas del proyecto. Necesario.

    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;El [parámetro opcional](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) siguiente:


&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **array, boolean, null, number, object, string** o **undefined**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto [AsyncResult](../../reference/shared/asyncresult.md) sin sufrir modificaciones. Opcional.</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Por ejemplo, puede pasar el argumento _asyncContext_ usando el formato `{asyncContext: 'Some text'}` o `{asyncContext: <object>}`.

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **function**

&nbsp;&nbsp;&nbsp;&nbsp;Una función que se invoca cuando se devuelve la llamada al método, cuyo único parámetro es del tipo [AsyncResult](../../reference/shared/asyncresult.md). Opcional.


## Valor de devolución de llamada

Cuando la función _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el parámetro de la función de devolución de llamada.

En el caso del método **getTaskByIndexAsync**, el objeto [AsyncResult](../../reference/shared/asyncresult.md) devuelto contiene las siguientes propiedades.


|**Nombre**|**Descripción**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Datos pasados en el parámetro opcional _asyncContext_, si se usó el parámetro.|
|[error](../../reference/shared/asyncresult.error.md)|Información sobre el error si la propiedad **status** es igual a **failed**.|
|[estado](../../reference/shared/asyncresult.status.md)|Estado **succeeded** o **failed** de la llamada asincrónica.|
|[value](../../reference/shared/asyncresult.value.md)|GUID de la tarea como una **string**.|

## Comentarios

Para obtener el índice máximo de la colección de tareas del proyecto, use el método [getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md). La tarea de índice 0 representa la tarea de resumen del proyecto.


## Ejemplo

En el siguiente ejemplo de código se llama a [getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md) para obtener el índice máximo de la colección de tareas del proyecto y luego se llama a **getTaskByIndexAsync** para obtener el GUID de cada tarea.

En el ejemplo se asume que el complemento tiene una referencia a la biblioteca de jQuery y que los controles de la siguiente página se definen en el div de contenido del cuerpo de la página.




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";
    var taskGuids = [];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#get-info').click(getTaskInfo);
        });
    };

    // Get the maximum task index, and then get the task GUIDs.
    function getTaskInfo() {
        getMaxTaskIndex().then(
            function (data) {
                getTaskGuids(data);
            }
        );
    }

    // Get the maximum index of the tasks for the current project.
    function getMaxTaskIndex() {
        var defer = $.Deferred();
        Office.context.document.getMaxTaskIndexAsync(
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

    // Get each task GUID, and then display the GUIDs in the add-in.
    function getTaskGuids(maxTaskIndex) {
        var defer = $.Deferred();
        for (var i = 0; i <= maxTaskIndex; i++) {
            getTaskGuid(i);
        }
        return defer.promise();
        function getTaskGuid(index) {
            Office.context.document.getTaskByIndexAsync(index,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        taskGuids.push(result.value);
                        if (index == maxTaskIndex) {
                            defer.resolve();
                            $('#message').html(taskGuids.toString());
                        }
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
    }
    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
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
|**Disponible en los conjuntos de requisitos **||
|**Nivel de permisos mínimo**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad

|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Agregado|

## Vea también



#### Otros recursos


[getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md)
[Objeto AsyncResult](../../reference/shared/asyncresult.md)
[Objeto ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
