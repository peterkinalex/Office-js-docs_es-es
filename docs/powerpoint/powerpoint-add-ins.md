
# Creación de complementos de contenido y panel de tareas para PowerPoint

Los ejemplos de código en este artículo muestran algunas tareas básicas para desarrollar complementos de contenido de PowerPoint. Para mostrar información, estos ejemplos dependen de la función  `app.showNotification`, que está incluida en las plantillas de proyecto Visual StudioComplementos de Office. Si no usa Visual Studio para desarrollar su complemento, deberá reemplazar la función  `showNotification` con su propio código. Varios de estos ejemplos también dependen de este objeto `globals` que se declara fuera del ámbito de estas funciones: `var globals = {activeViewHandler:0, firstSlideId:0};`

Estos ejemplos de código exigen que su proyecto [haga referencia a la biblioteca de Office.js v1.1 o posterior](../../docs/develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).


## Detecte la vista activa de la presentación y maneje el evento ActiveViewChanged

La función  `getFileView` llama al método [Document.getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md) para que regrese independientemente de que la vista actual de la presentación sea la vista "editar" (cualquiera de las vistas en las que puede editar diapositivas, como **Normal** o **Vista de esquema**) o "lectura" ( **Presentación de diapositivas** o **Vista de lectura**).


```js
function getFileView() {
    //Gets whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });
}
```

La función  `registerActiveViewChanged` llama al método [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) para registrar un controlador para evento [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md). Después de ejecutar esta función, cuando cambie la vista de la presentación, la notificación  `app.showNotification` mostrará el modo de vista activa ("leer" o "editar").




```js
function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```


## Obtenga la dirección URL de la presentación

La función `getFileUrl` llama al método [Document.getFileProperties](../../reference/shared/document.getfilepropertiesasync.md) para obtener la dirección URL del archivo de presentación.


```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```


## Navegar a un diapositiva particular en la presentación

La función  `getSelectedRange` llama al método [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) para obtener un objeto JSON devuelto por `asyncResult.value`, que contiene una variedad de "diapositivas" con nombre que incluyen los identificadores, títulos e índices del rango seleccionado de diapositivas (o solo la diapositiva actual). También guarda el identificador de la primera diapositiva en el rango seleccionado en una variable global.


```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

La función  `goToFirstSlide` llama al método [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) para ir a la identificación de la primera diapositiva almacenada por la función `getSelectedRange` anterior.




```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## Navegar entre diapositivas en la presentación

La función  `goToSlideByIndex` llama al método **Document.goToByIdAsync** para navegar hasta la siguiente diapositiva en la presentación.


```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```




## Recursos adicionales

- [Procedimiento para guardar la configuración y el estado de los complementos por cada documento de los complementos de panel de tareas y de contenido](../../docs/develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [Leer y escribir datos en la selección activa de un documento o una hoja de cálculo](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [Procedimiento para obtener el documento completo de un complemento para PowerPoint o Word](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [Usar los temas del documento en los complementos de PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
