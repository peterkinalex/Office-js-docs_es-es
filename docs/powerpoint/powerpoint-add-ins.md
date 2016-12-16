# <a name="powerpoint-add-ins"></a>Complementos de PowerPoint

Puede utilizar complementos de PowerPoint para compilar soluciones más atractivas para las presentaciones de sus usuarios en plataformas como Windows, iOS, Office Online y Mac. Puede crear uno de los dos tipos de complementos:

- Utilice **complementos de contenido** para agregar contenido dinámico de HTML5 a sus presentaciones. Por ejemplo, consulte el complemento de [diagramas LucidChart para PowerPoint](https://store.office.com/en-us/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false), que puede utilizar para insertar un diagrama interactivo de LucidChart en la baraja.
- Utilice los **complementos del panel de tareas** para introducir información de referencia o insertar datos en la diapositiva mediante un servicio. Por ejemplo, consulte el complemento [imágenes Shutterstock](https://store.office.com/en-us/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false), que puede utilizar para agregar fotos profesionales a su presentación. 

>**Nota:** Al generar el complemento, si va a [publicar](../publish/publish.md) el complemento en la Tienda Office, asegúrese de que se ajustan a la [directivas de validación de la Tienda Office](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Por ejemplo, para superar la validación, el complemento debe funcionar en todas las plataformas que sean compatibles con los métodos especificados en el elemento Requirements del manifiesto (vea la [sección 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3)).

## <a name="powerpoint-add-in-scenarios"></a>Escenarios de complementos de PowerPoint

Los ejemplos de código del artículo muestran algunas tareas básicas para desarrollar complementos de contenido para PowerPoint. 

Para mostrar información, estos ejemplos dependen de la función `app.showNotification`, que se incluye en las plantillas de proyecto de los complementos de Office de Visual Studio. Si no utiliza Visual Studio para desarrollar su complemento, deberá reemplazar la función `showNotification` con su propio código. Algunos de estos ejemplos también dependen de este objeto `globals` que se declara fuera del ámbito de estas funciones: `var globals = {activeViewHandler:0, firstSlideId:0};`

Estos ejemplos de código exigen que su proyecto [haga referencia a la biblioteca de Office.js v1.1 o posterior](../../docs/develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>Detecte la vista activa de la presentación y maneje el evento ActiveViewChanged

Si va a crear un complemento de contenido, tendrá que obtener la vista activa de la presentación y manejar el evento ActiveViewChanged, como parte de su controlador Office.Initialize.


- La función  `getActiveFileView` llama al método [Document.getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md) para que regrese independientemente de que la vista actual de la presentación sea la vista "editar" (cualquiera de las vistas en las que puede editar diapositivas, como **Normal** o **Vista de esquema**) o "lectura" ( **Presentación de diapositivas** o **Vista de lectura**).


- La función `registerActiveViewChanged` llama al método [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) para registrar un controlador para el evento [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md). 
> Nota: En PowerPoint Online, el evento [ Document.ActiveViewChanged ](../../reference/shared/document.activeviewchanged.md) no se iniciará nunca, ya que el modo Presentación con diapositivas se trata como una nueva sesión. En este caso, el complemento debe capturar la vista activa en carga, como se indica a continuación.



```js

//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    var currentView = getActiveFileView();

    //register for the active view changed handler
    registerActiveViewChanged();

    //render the content based off of the currentView
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });

}


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
    

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>Navegar a una diapositiva particular en la presentación

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


## <a name="navigate-between-slides-in-the-presentation"></a>Navegar entre diapositivas en la presentación

La función `goToSlideByIndex` llama al método **Document.goToByIdAsync** para navegar hasta la siguiente diapositiva en la presentación.


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

## <a name="get-the-url-of-the-presentation"></a>Obtenga la dirección URL de la presentación

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



## <a name="additional-resources"></a>Recursos adicionales
- [Ejemplos de código de PowerPoint](https://dev.office.com/code-samples#?filters=powerpoint)

- [Procedimiento para guardar la configuración y el estado de los complementos por cada documento de los complementos de panel de tareas y de contenido](../../docs/develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [Leer y escribir datos en la selección activa de un documento o una hoja de cálculo](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [Obtener el documento completo de un complemento para PowerPoint o Word](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [Usar temas de documentos en los complementos para PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
