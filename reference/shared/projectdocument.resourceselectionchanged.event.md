

# <a name="projectdocument.resourceselectionchanged-event"></a>Evento ProjectDocument.ResourceSelectionChanged
Se produce cuando la selección de recursos cambia en el proyecto activo.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selección|
|**Agregado en**|1.0|

```js
Office.EventType.ResourceSelectionChanged
```


## <a name="remarks"></a>Observaciones

 **ResourceSelectionChanged** es una constante de enumeración de [EventType](../../reference/shared/eventtype-enumeration.md) que se puede usar en los métodos [ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) y [ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md) para agregar o quitar un controlador para el evento.


## <a name="example"></a>Ejemplo

En el ejemplo de código siguiente se agrega un controlador para el evento **ResourceSelectionChanged**. Al cambiar la selección del recurso en el documento, se obtiene el GUID del recurso seleccionado.

En el ejemplo se presupone que el complemento tiene una referencia a la biblioteca de jQuery y que el siguiente control de página se define en div de contenido en el cuerpo de la página.




```HTML
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
        });
    };

    // Get the GUID of the selected resource and display it in the add-in.
    function getResourceGuid() {
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html(result.value);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```

Para obtener un ejemplo de código completo en el que se muestre cómo usar un controlador de eventos **ResourceSelectionChanged** en un complemento de Project, consulte [Create your first task pane app for Project 2013 by using a text editor](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md) (Crear su primer complemento de panel de tareas para Project 2013 con un editor de texto).


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este evento es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este evento.

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


[Create your first task pane add-in for Project 2013 by using a text editor (Crear su primer complemento de panel de tareas para Project 2013 con un editor de texto)](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
[Enumeración EventType](../../reference/shared/eventtype-enumeration.md)
[Método ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)
[Método ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md)
[Objeto ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
