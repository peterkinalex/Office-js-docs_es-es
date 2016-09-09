

# Evento ProjectDocument.ViewSelectionChanged
Ocurre cuando la visualización activa cambia en el proyecto activo.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selección|
|**Agregado en**|1,0|

```js
Office.EventType.ViewSelectionChanged
```


## Observaciones

 **ViewSelectionChanged** es una constante de enumeración de [EventType](../../reference/shared/eventtype-enumeration.md) que se puede usar en los métodos [ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) y [ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md) para agregar o quitar un controlador para el evento.


## Ejemplo

En el ejemplo de código siguiente se agrega un controlador para el evento **ViewSelectionChanged**. Cuando se cambia la vista activa, obtiene el nombre y el tipo de la vista activa.

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
                Office.EventType.ViewSelectionChanged,
                getActiveView);
            getActiveView();
        });
    };

    // Get the name and type of the active view and display it in the add-in.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, result.value.viewType);
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

Para ver un ejemplo en el que se muestre como usar un controlador de eventos **ViewSelectionChanged** en un complemento de Project, consulte [Create your first task pane add-in for Project 2013 by using a text editor](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md) (Crear su primer complemento de panel de tareas para Project 2013 con un editor de texto).


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este evento es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este evento.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office para escritorio de Windows**|**Office Online (en el explorador)**|
|:-----|:-----|:-----|
|**Project**|v||

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **||
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



|**Versión**|**Cambios**|
|:-----|:-----|
|1,0|Agregado|

## Vea también



#### Otros recursos


[Create your first task pane add-in for Project 2013 by using a text editor (Crear su primer complemento de panel de tareas para Project 2013 con un editor de texto)](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
[Enumeración EventType](../../reference/shared/eventtype-enumeration.md)
[Método ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)
[Método ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md)
[Objeto ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)

