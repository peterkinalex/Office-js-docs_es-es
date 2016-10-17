
# <a name="document-object"></a>Objeto Document
Una clase abstracta que representa el documento con el que interactúa el complemento.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project y Word|
|**Agregado en**|1.0|
|**Modificado por última vez en**|1.1|

```
Office.context.document
```


## <a name="members"></a>Miembros


**Propiedades**


|**Nombre**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|
|[bindings](../../reference/shared/document.bindings.md)|Obtiene un objeto que proporciona acceso a los enlaces que se han definido en el documento.|En 1.1 se agregó compatibilidad para los complementos de contenido para Access.|
|[customXmlParts](../../reference/shared/document.customxmlparts.md)|Obtiene un objeto que representa los elementos XML personalizados del documento.||
|[mode](../../reference/shared/document.mode.md)|Obtiene el modo en el que se encuentra el documento.|En 1.1 se agregó compatibilidad para los complementos de contenido para Access.|
|[settings](../../reference/shared/document.settings.md)|Obtiene un objeto que representa la configuración personalizada que se ha guardado del complemento de contenido o del panel de tareas para el documento actual.|En 1.1 se agregó compatibilidad para los complementos de contenido para Access.|
|[url](../../reference/shared/document.url.md)|Obtiene la dirección URL del documento que se encuentra abierto actualmente en la aplicación host.|En 1.1 se agregó compatibilidad para los complementos de contenido para Access.|

**Métodos**


|**Nombre**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|
|[addHandlerAsync](../../reference/shared/document.addhandlerasync.md)|Agrega un controlador de eventos para un evento del objeto **Document**.||
|[getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md)|Devuelve la vista actual de la presentación.|En 1.1 se agregó compatibilidad para admitir [los complementos para PowerPoint](../../docs/powerpoint/powerpoint-add-ins.md).|
|[getFileAsync](../../reference/shared/document.getfileasync.md)|Devuelve el archivo de documento entero en segmentos de hasta 4194304 bytes (4MB).|En 1.1, se agregó compatibilidad para obtener el archivo como PDF en los complementos para PowerPoint y Word.|
|[getFilePropertiesAsync](../../reference/shared/document.getfilepropertiesasync.md)|Obtiene las propiedades de archivo del documento actual. En esta versión solo puede obtener la dirección URL del documento.|En 1.1 se agregó para obtener la dirección URL del documento en los complementos para Excel, Word y PowerPoint.|
|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|Lee los datos incluidos en la selección actual del documento.|En 1.1 se agregó compatibilidad para obtener el identificador, el título y el índice del intervalo de diapositivas seleccionado en los complementos para PowerPoint.|
|[goToByIdAsync](../../reference/shared/document.gotobyidasync.md)|Va al objeto o la ubicación que se haya especificado en el documento.|En 1.1 se agregó la compatibilidad para navegar por el documento en los complementos para Excel y PowerPoint.|
|[removeHandlerAsync](../../reference/shared/document.removehandlerasync.md)|Quita un controlador de eventos para un evento del objeto **Document**.||
|[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|Escribe datos en la selección actual del documento.|En 1.1 se aumentó la compatibilidad para [establecer el formato de la tabla seleccionada al escribir datos en los complementos para Excel](../../docs/excel/format-tables-in-add-ins-for-excel.md).|

**Eventos**


|**Nombre**|**Descripción**|**Notas de compatibilidad**||
|:-----|:-----|:-----|:-----|
|[ActiveViewChanged](../../reference/shared/document.activeviewchanged.md)|Se produce cuando el usuario cambia la vista actual del documento.|En 1.1 se agregó para admitir los complementos para PowerPoint.||
|[SelectionChanged](../../reference/shared/document.selectionchanged.event.md)|Se produce al cambiar la selección en el documento.|||

## <a name="remarks"></a>Comentarios

No puede crear instancias del objeto **Document** directamente en el script. Si desea llamar a miembros del objeto **Document** para que interactúen con la hoja de cálculo o el documento actual, use `Office.context.document` en el script.


## <a name="example"></a>Ejemplo

En el ejemplo siguiente se usa el método **getSelectedDataAsync** del objeto **Document** para recuperar la selección actual del usuario como texto y, a continuación, mostrarla en la página del complemento.


```js

// Display the user's current selection.
function showSelection() {
    Office.context.document.getSelectedDataAsync(
        "text",                        // coercionType
        {valueFormat: "unformatted",   // valueFormat
        filterType: "all"},            // filterType
        function (result) {            // callback
            var dataValue; 
            dataValue = result.value;
            write('Selected data is: ' + dataValue);
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>Detalles de compatibilidad


La compatibilidad para cada miembro de API del objeto **Document** difiere entre aplicaciones host de Office. Consulte la sección "Detalles de compatibilidad" del tema de cada miembro para obtener información de compatibilidad de host.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Agregado en**|1.0|
|**Modificado por última vez en**|1.1|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|
