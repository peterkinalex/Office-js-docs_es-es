

# Objeto ProjectDocument
Clase abstracta que representa el documento del proyecto (el proyecto activo) con el que interactúa el complemento de Office.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Agregado en**|1,0|

```js
Office.context.document
```


## Miembros


**Métodos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[Método addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)|Agrega de forma asíncrona un controlador de evento para un evento en un objeto **ProjectDocument**.|
|[Método getMaxResourceIndexAsync](../../reference/shared/projectdocument.getmaxresourceindexasync.md)|Obtiene de forma asincrónica el índice máximo de la colección de recursos del proyecto actual.|
|[Método getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md)|Obtiene de forma asincrónica el índice máximo de la colección de tareas del proyecto actual.|
|[Método getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)|Obtiene de forma asíncrona el valor del campo especificado del proyecto activo.|
|[Método getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md)|Obtiene de forma asíncrona el GUID del recurso que tiene el índice especificado en la colección de recursos.|
|[Método getResourceFieldAsync](../../reference/shared/projectdocument.getresourcefieldasync.md)|Obtiene de forma asincrónica el valor del campo indicado para el recurso que se ha especificado.|
|[Método getSelectedDataAsync](../../reference/shared/projectdocument.getselecteddataasync.md)|Obtiene de forma asincrónica los datos que contiene la selección actual de una o varias celdas del diagrama de Gantt.|
|[Método getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedresourceasync.md)|Obtiene de forma asincrónica el identificador único global (GUID, Globally Unique Identifier) del recurso seleccionado.|
|[Método getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)|Obtiene de forma asincrónica el GUID de la tarea seleccionada.|
|[Método getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)|Obtiene de forma asincrónica el nombre y el tipo de vista de la vista activa.|
|[Método getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md)|Obtiene de forma asincrónica el nombre de la tarea, los recursos que se han asignado a la misma y el identificador correspondiente de la lista sincronizada de tareas de SharePoint.|
|[Método getTaskByIndexAsync](../../reference/shared/projectdocument.gettaskbyindexasync.md)|Obtiene de forma asincrónica el GUID de la tarea que tiene el índice especificado en la colección de tareas.|
|[Método getTaskFieldAsync](../../reference/shared/projectdocument.gettaskfieldasync.md)|Obtiene de forma asíncrona el valor del campo especificado para la tarea especificada.|
|[Método getWSSUrlAsync](../../reference/shared/projectdocument.getwssurlasync.md)|Obtiene de forma asincrónica la dirección URL de la lista sincronizada de tareas de SharePoint.|
|[Método removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md)|Elimina de forma asíncrona un controlador de evento para un evento de cambio en un objeto **ProjectDocument**.|
|[Método setResourceFieldAsync](../../reference/shared/projectdocument.setresourcefieldasync.md)|Establece de forma asincrónica el valor del campo especificado para el recurso concreto.|
|[Método setTaskFieldAsync](../../reference/shared/projectdocument.settaskfieldasync.md)|Establece de forma asincrónica el valor del campo especificado para la tarea determinada.|

**Eventos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[ResourceSelectionChanged event](../../reference/shared/projectdocument.resourceselectionchanged.event.md)|Ocurre cuando la selección de recursos cambia en el proyecto activo.|
|[TaskSelectionChanged event](../../reference/shared/projectdocument.taskselectionchanged.event.md)|Se genera cuando cambia la selección de tareas en el proyecto activo.|
|[ViewSelectionChanged event](../../reference/shared/projectdocument.viewselectionchanged.event.md)|Ocurre cuando la visualización activa cambia en el proyecto activo.|

## Comentarios

No llame directamente o cree instancias del objeto **ProjectDocument** en el script.


## Ejemplo

En el siguiente ejemplo se inicializa el complemento y después se obtienen propiedades del objeto [Document](../../reference/shared/document.md) que están disponibles en el contexto de un documento de Project. Un documento de Project es el proyecto activo y abierto. Para tener acceso a los miembros del objeto **ProjectDocument**, use el objeto **Office.context.document** tal como se muestra en los ejemplos de código para los eventos y métodos **ProjectDocument**.

En el ejemplo se presupone que el complemento tiene una referencia a la biblioteca de jQuery y que el siguiente control de página se define en div de contenido en el cuerpo de la página:




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Get information about the document.
            showDocumentProperties();
        });
    };

    // Get the document mode and the URL of the active project.
    function showDocumentProperties() {
        var output = String.format(
            'The document mode is {0}.<br/>The URL of the active project is {1}.',
            Office.context.document.mode,
            Office.context.document.url);
        $('#message').html(output);
    }
})();
```


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|
|:-----|:-----|:-----|
|**Project**|v||

|||
|:-----|:-----|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad


|**Versión**|**Cambios**|
|:-----|:-----|
|1,0|Agregado|

## Vea también



#### Otros recursos


[Complementos de panel de tareas para Project](../../docs/project/project-add-ins.md)
[Objeto Document](../../reference/shared/document.md)

