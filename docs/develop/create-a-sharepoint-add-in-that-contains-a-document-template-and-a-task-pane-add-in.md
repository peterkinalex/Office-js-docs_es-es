
# Crear un complemento de SharePoint que contenga una plantilla de documento y un complemento de panel de tareas


Puede crear un Complemento de SharePoint que incluya una plantilla de documento (por ejemplo, un informe de gastos). El documento puede incluir un complemento de panel de tareas que interactúe con datos de SharePoint. Por ejemplo, los usuarios pueden rellenar los campos de una factura con datos de los servicios de conectividad empresarial (BCS) o crear un informe de gastos seleccionando una categoría de gastos de una lista de SharePoint.

Este tutorial muestra cómo crear un Complemento de SharePoint que incluye un libro de Excel. El libro de Excel contiene un complemento de panel de tareas que usa la interfaz REST suministrada por SharePoint 2013 para rellenar un cuadro de lista desplegable con la fecha de SharePoint en el complemento de panel de tareas.


## Requisitos previos


Instale los siguientes componentes antes de empezar:




- Un entorno de desarrollo de SharePoint:
    
      - To develop SharePoint Add-ins that target SharePoint in Office 365, see [How to: Set up an environment for developing SharePoint Add-ins on Office 365](http://msdn.microsoft.com/en-us/library/office/apps/fp161179%28v=office.15%29).
    
  - Para desarrollar Complementos de SharePoint orientados a una instalación local de SharePoint, vea [Procedimiento para preparar un entorno de desarrollo local para complementos para SharePoint](http://msdn.microsoft.com/en-us/library/office/apps/fp179923%28v=office.15%29).
    
- [Visual Studio 2015 y Microsoft Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs)
    
- Excel 2013 o una cuenta de Office 365.
    

## Crear un proyecto de Complemento de SharePoint en Visual Studio



1. Inicie Visual Studio.
    
2. En la barra de menú, elija  **Archivo**,  **Nuevo** y **Proyecto**.
    
    Se abre el cuadro de diálogo **Nuevo proyecto**.
    
3. En el panel de plantillas, debajo del nodo para el lenguaje que quiera usar, expanda  **Office SharePoint** y elija **Complementos de Office**.
    
4. En la lista de tipos de proyecto, elija  **Complemento de SharePoint**, asigne al proyecto el nombre OfficeEnabledAddin y, a continuación, elija el botón **Aceptar**.
    
    Aparecerá el cuadro de diálogo **Nuevo complemento de SharePoint**.
    
5. En la lista desplegable de  **¿Qué sitio de SharePoint desea usar para depurar el complemento?**, elija o escriba la dirección URL de un sitio de SharePoint.
    
6. En la lista desplegable  **¿Cómo desea hospedar el Complemento de SharePoint?**, elija  **Hospedaje en SharePoint** y, a continuación, elija **Siguiente**.
    
     >**Nota** Este escenario funciona únicamente con las opciones de hospedaje en SharePoint o en proveedores que se presentan en la lista desplegable **¿Cómo desea hospedar el complemento de SharePoint?**
7. En la siguiente página, seleccione  **SharePoint 2013** y, a continuación, elija el botón **Finalizar** para cerrar el cuadro de diálogo.
    

## Agregar un elemento de complemento de panel de tareas.


Después, agregue un Office Add-in al proyecto. Puede agregar cualquier tipo de complemento que desee. En este tutorial agregaremos un complemento de panel de tareas.


1. En el  **Explorador de soluciones**, elija el nodo de proyecto  **OfficeEnabledAddin**.
    
2. En el menú  **Proyecto**, elija  **Agregar nuevo elemento**.
    
3. En el cuadro de diálogo  **Agregar nuevo elemento**, seleccione  **Office/SharePoint** y, a continuación, elija **Complemento de Office**.
    
4. Asigne al complemento de panel de tareas el nombre MyTaskPaneAddin y seleccione el botón **Agregar**.
    
    Se abrirá el cuadro de diálogo **Crear un complemento para Office**.
    
5. En el cuadro de diálogo  **Crear complemento de Office**, seleccione  **Panel de tareas** y, a continuación, elija **Siguiente**. En la siguiente página, desactive las casillas  **Word** y **PowerPoint** y luego elija **Siguiente**.
    
6. En la página  **¿Desea que el Office Add-in aparezca en un documento nuevo o en uno existente?**, elija  **Crear un nuevo documento e insertar mi complemento** y después elija **Finalizar**.
    
    Visual Studio agrega una biblioteca de documentos y una plantilla de libro para la biblioteca. El libro contiene un complemento de panel de tareas.
    

## Agregar una biblioteca de documentos


En este procedimiento, agregará una biblioteca de documentos y hará que el libro sea la plantilla predeterminada de la biblioteca de documentos.


1. En el  **Explorador de soluciones**, elija el nodo de proyecto  **OfficeEnabledAddin**.
    
2. En el menú  **Proyecto**, elija  **Agregar nuevo elemento**.
    
3. En el cuadro de diálogo  **Agregar nuevo elemento**, seleccione  **Office/SharePoint**, elija  **Lista**, asigne el nombre MyDocumentLibrary a la lista y seleccione el botón **Agregar**.
    
4. En el  **Asistente para la personalización de SharePoint**, seleccione la opción  **Crear una plantilla de lista personalizada y una instancia de lista de ella**.
    
5. En la lista desplegable debajo de esta opción, seleccione  **Biblioteca de documentos** y, después, elija el botón **Siguiente**.
    
6. En la página  **Elija una plantilla para esta biblioteca de documentos. Los documentos que los usuarios creen en esta biblioteca se basarán en esa plantilla**, elija  **Usar el siguiente documento como la plantilla para esta biblioteca** y, después, elija el botón **Explorar**.
    
7. En el cuadro de diálogo  **Abrir**, abra la carpeta  **OfficeDocuments**, elija el archivo  **MyTaskPaneApp.xlsx**, seleccione el botón  **Abrir**, elija el botón  **Finalizar** y cierre el diseñador de listas.
    
8. En el  **Explorador de soluciones**, elija el nodo de proyecto  **OfficeEnabledAddin**.
    
9. En el menú  **Vista**, elija  **Ventana Propiedades**.
    
10. En el  **Explorador de soluciones**, elija el archivo  **AppManifest.xml**.
    
11. Elija  **Ver** y **Diseñador**.
    
12. En el diseñador de manifiestos, establezca el valor de la  **Página de inicio** en~appWebUrl/Lists/MyDocumentLibrary. Esto se convertirá en un valor de OfficeEnabledAddin/Lists/MyDocumentLibrary.
    
     >**Nota** Esta dirección URL hace referencia a la biblioteca de documentos. Debe usar el token ~appWebUrl al principio de cualquier dirección URL del manifiesto de complemento de Office que hace referencia a elementos dentro de la Web de complemento. Para obtener más información sobre tokens de direcciones URL en un proyecto de complemento de SharePoint, consulte [Cadenas y tokens de direcciones URL en complementos de SharePoint](http://msdn.microsoft.com/library/800ec8cd-a448-46bc-b41e-d4030eeb4048%28Office.15%29.aspx).
13. Cierre el diseñador de manifiestos para guardar el cambio.
    

## Consumir datos de SharePoint en el panel de tareas


En este procedimiento, mostrará una lista de usuarios del sitio mediante el uso de la interfaz de transferencia de estado representacional (REST) que proporciona SharePoint 2013.

En este ejemplo, los datos de la lista de SharePoint solo se muestran, pero puede usar este tipo de datos como parte de un complemento de aprobación de documentos. Cuando un usuario elige un nombre de la lista, su código establece el valor de la columna del revisor en una lista de seguimiento de documentos. Un flujo de trabajo asociado a esa lista podría enviar una notificación de revisión a dicho usuario. Como alternativa, puede guardar el nombre seleccionado en la configuración del documento. A continuación, cuando un usuario abra el documento, podrá mostrar controles en el complemento de panel de tareas solo si el usuario actual y el usuario guardado en la configuración del documento coinciden. Para más información, vea los temas siguientes:


- [Procedimiento para realizar operaciones básicas con extremos REST de SharePoint 2013](http://msdn.microsoft.com/library/e3000415-50a0-426e-b304-b7de18f2f7d9%28Office.15%29.aspx)
    
- [Completar operaciones básicas con código de biblioteca de JavaScript en SharePoint 2013](http://msdn.microsoft.com/library/29089af8-dbc0-49b7-a1a0-9e311f49c826%28Office.15%29.aspx)
    
- [Conservación de la configuración y del estado de los complementos](../../docs/develop/persisting-add-in-state-and-settings.md)
    

1. En el  **Explorador de soluciones**, expanda las carpetas  **MyTaskPaneAddin** y **Inicio** y seleccione el archivo **Home.html**.
    
    El archivo Home.html se abre en el editor de código.
    
2. Agregue el siguiente HTML bajo el botón  `get-data-from-selection`.
    
```HTML
  <p>Select Reviewer:</p> <select class="select" id="select-reviewer" name="D1"> </select>
```

3. Seleccione el archivo  **Home.js** para que se abra en el editor de código.
    
4. Agregue las siguientes declaraciones a la parte superior del archivo Home.js.
    
```js
  var appWebURL; var web;
```

5. Reemplace la función  `Initialize` con el siguiente código.
    
    Este código realiza las siguientes tareas:
    
      - Carga los archivos SP.Runtime.js y SP.js con la función  `getScript` en jQuery. Después de cargar los archivos, el programa tiene acceso al modelo de objetos JavaScript para SharePoint.
    
  - Carga el objeto del sitio web actual.
    
  - Llama a una función que obtiene todos los usuarios del sitio. Agregará el código de dicha función en el siguiente paso.
    



```js
   // The initialize function must be run each time a new page is loaded Office.initialize = function (reason) { $(document).ready(function () { app.initialize(); var scriptbase = "/_layouts/15/"; $.getScript(scriptbase + "SP.Runtime.js", function () { $.getScript(scriptbase + "SP.js", function () { getAppWeb(function () { getSPUsers(populateUsersDropDown); }); }); }); function getAppWeb(functionToExecuteOnReady) { var context = SP.ClientContext.get_current(); web = context.get_web(); context.load(web); context.executeQueryAsync(onSuccess, onFailure); function onSuccess() { appWebURL = web.get_url(); functionToExecuteOnReady(); } function onFailure(sender, args) { app.initialize(); app.showNotification("Failed to connect to SharePoint. Error: " + args.get_message()); } } $('#get-data-from-selection').click(getDataFromSelection); }); };
```

6. Agregue el siguiente código en la parte inferior del archivo Home.js.
    
    Este código obtiene una lista de usuarios de sitios web mediante la interfaz REST que proporciona SharePoint 2013. A continuación, este código rellena una lista desplegable con los nombres y los identificadores de cada uno de los usuarios.
    


```js
  function getSPUsers(functionToExecuteOnReady) { var url = appWebURL + "/../_api/web/siteUsers"; jQuery.ajax({ url: url, type: "GET", headers: { "ACCEPT": "application/json;odata=verbose" }, success: onSuccess, error: onFailure }); function onSuccess(data) { var results = data.d.results; functionToExecuteOnReady(results); } function onFailure(jaXHR, textStatus, errorThrown) { var error = textStatus + " " + errorThrown; app.showNotification(error); } } function populateUsersDropDown(results) { for (var i = 0; i < results.length; i++) { var IDTemp = results[i].Id; $('#select-reviewer').append("<option value='" + IDTemp + "'>" + results[i].Title + "</option>"); } }
```

7. En  **Explorador de soluciones**, abra el menú contextual del archivo  **AppManifest.xml** y, después, elija **Diseñador de vistas**.
    
8. En el diseñador, elija la página  **Permisos**.
    
9. Desde la lista desplegable de la columna  **Ámbito**, elija el elemento  **Web**.
    
10. Desde la lista desplegable de la columna  **Permisos**, elija el elemento  **Lectura** item.
    

## Depurar el complemento del panel de tareas


Puede depurar el complemento del panel de tareas abriendo el documento o iniciando el Complemento de SharePoint y, luego, abriendo un documento de la biblioteca.


### Depuración del complemento del panel de tareas iniciando el documento




 >**Nota**  Debido a que este procedimiento abre Excel, solo funciona cuando Office está instalado en el sistema. De lo contrario, obtendrá este error: "La aplicación asociada a este tipo de proyecto no está instalada en el equipo".


1. Abra el archivo Home.js en el editor de código y, después, establezca un punto de interrupción junto al método  `getDataFromSelection`.
    
2. En el  **Explorador de soluciones**, seleccione el nodo de proyecto  **OfficeEnabledApp**.
    
3. En el menú  **Vista**, elija  **Ventana Propiedades**.
    
4. En la ventana Propiedades, vaya a la lista desplegable  **Acción de inicio** y elija el elemento **Cliente de escritorio Office**. Cuando lo haga, aparecerá una propiedad nueva,  **Documento de inicio**.
    
5. En la lista desplegable  **Documento de inicio**, elija el elemento  **OfficeDocuments\TaskPaneApp.xlsx**.
    
6. En el menú  **Depurar**, elija  **Iniciar depuración**.
    
    Este valor hace que el libro del complemento de panel de tareas aparezca cuando se ejecuta el complemento. El libro se abre y aparece el complemento de panel de tareas.
    
7. En el complemento del panel de tareas, elija la lista desplegable  **Seleccionar revisor** para ver una lista de usuarios de SharePoint.
    
8. En el libro de Excel, seleccione una celda cualquiera.
    
9. En el complemento de panel de tareas, elija el botón  **Obtener datos de la selección**.
    
    La ejecución se detiene en el punto de interrupción que estableció junto al método `getDataFromSelection`.
    

### Depuración del complemento de panel de tareas iniciando SharePoint




 >
  **Nota**  Este procedimiento abre Excel Online y solo funciona si se dispone de una cuenta de Office 365. Vea [Procedimiento para preparar un entorno de desarrollo de complemento para SharePoint en Office 365](http://msdn.microsoft.com/en-us/library/office/apps/fp161179%28v=office.15%29).


1. Abra el archivo Home.js en el editor de código y, después, defina un punto de interrupción junto al método  `getDataFromSelection`.
    
2. En el  **Explorador de soluciones**, seleccione el nodo de proyecto  **OfficeEnabledApp**.
    
3. En el menú  **Vista**, elija  **Ventana Propiedades**.
    
4. En la ventana Propiedades, vaya a la lista desplegable  **Acción de inicio** y elija el elemento **Internet Explorer**.
    
5. En el menú  **Depurar**, elija  **Iniciar depuración**.
    
    Visual Studio abre SharePoint y muestra la biblioteca **MyDocumentLibrary**.
    
6. En SharePoint, vaya a la pestaña  **Archivos** y seleccione **Nuevo documento**. 
    
7. Navegue hasta el libro del proyecto, MyTaskPaneApp.xlsx.
    
    El libro se abre y aparece el complemento de panel de tareas.
    
8. Asegúrese de que la depuración de scripts esté habilitada en el explorador. Para habilitar la depuración de scripts en Internet Explorer, abra el cuadro de diálogo  **Opciones de Internet**, elija la pestaña  **Opciones avanzadas** y desactive las casillas **Deshabilitar la depuración de scripts (Internet Explorer)** y **Deshabilitar la depuración de scripts (otros)**.
    
9. En Visual Studio, en el menú  **Depurar**, elija  **Asociar al proceso**.
    
10. En el cuadro de diálogo  **Asociar al proceso**, elija todos los procesos  **iexplore.exe** disponibles y después seleccione el botón **Asociar**.
    
11. En el complemento de panel de tareas, elija la lista desplegable  **Seleccionar revisor** para ver una lista de usuarios de SharePoint.
    
    Los datos de esta lista se obtienen de SharePoint mediante una llamada a REST.
    
12. En el libro de Excel, elija una celda cualquiera.
    
13. En el complemento de panel de tareas, elija el botón  **Obtener datos de la selección**.
    
    La ejecución se detiene en el punto de interrupción que estableció junto al método `getDataFromSelection`.
    
     >**Nota** Si el libro no contiene datos y desea agregarlos, en la barra de herramientas del libro, elija **EDITAR LIBRO** y **Editar en Excel en línea**.

## Empaquetar y publicar el complemento


Cuando esté listo para empaquetar el complemento para publicarlo, abra el asistente para  **Publicar Complementos de SharePoint y Office**.


- En el **Explorador de soluciones**, abra el menú contextual para el proyecto de Complemento de SharePoint y, después, elija **Publicar**.
    
    Aparecerá el asistente **Publicar complementos para Office y SharePoint**. Para obtener más información, consulte [Publicar complementos para SharePoint con Visual Studio](http://msdn.microsoft.com/library/8137d0fa-52e2-4771-8639-60af80f693bb%28Office.15%29.aspx).
    

## Recursos adicionales


- [Directrices de diseño para complementos de Office](../../docs/design/add-in-design.md)
    
- [Ciclo de vida de desarrollo de complementos de Office](../../docs/design/add-in-development-lifecycle.md)
    
- [Publicar el complemento para Office](../publish/publish.md)
    
- [Información sobre la API de JavaScript para Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Manifiesto XML de complementos para Office](../../docs/overview/add-in-manifests.md)
    
- [Referencias de esquema y API de complementos de Office](../../reference/reference.md)
    
