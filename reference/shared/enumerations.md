
# Enumeraciones

Puede especificar un valor enumerado usando su nombre de enumeración completo (`Office.CoercionType.Text`) o su valor de texto correspondiente (`"text"`). Por ejemplo, la llamada de método siguiente usa nombres de enumeración:


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, {valueFormat:Office.ValueFormat.Unformatted, filterType:Office.FilterType.All},
   function (result) {
      if (result.status === Office.AsyncResultStatus.Success)
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {
         var err = result.error;
         write(err.name + ": " + err.message);
      }
   });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```


Y esta es la misma llamada usando los valores de texto de enumeración:




```js
Office.context.document.getSelectedDataAsync("text", {valueFormat:"unformatted", filterType:"all"},
   function (result) {
      if (result.status === "success")
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {
         var err = result.error;
         write(err.name + ": " + err.message);
      }
   });
```


## Referencia



|**Nombre**|**Definición**|
|:-----|:-----|
|[ActiveView](activeview-enumeration.md)|Especifica el estado de la vista activa del documento (por ejemplo, si el usuario puede editar o no el documento).|
|[AsyncResultStatus](asyncresultstatus-enumeration.md)|Especifica el resultado de una llamada asíncrona.|
|[AttachmentType](http://msdn.microsoft.com/library/83883a47-a937-4afb-a55e-e789057335c4%28Office.15%29.aspx)|Especifica el tipo de los datos adjuntos de un mensaje de correo electrónico o de una convocatoria de reunión. Outlook 2013 no admite esta enumeración.|
|[BindingType](bindingtype-enumeration.md)|Especifica el tipo de objeto de enlace que se debería devolver.|
|[BodyType](http://msdn.microsoft.com/library/31350fe6-4c42-4cbb-a5b2-4fb2d360fa11%28Office.15%29.aspx)|Especifica el tipo de texto para el cuerpo de una cita o un mensaje.|
|[CoercionType](coerciontype-enumeration.md)|Especifica cómo convertir los datos que el método invocado ha devuelto o definido.|
|[CustomXMLNodeType](customxmlnodetype-enumeration.md)|Especifica el tipo de nodo.|
|[DocumentMode](documentmode-enumeration.md)|Especifica si el documento de la aplicación asociada es de solo lectura o de lectura y escritura. |
|[EntityType](http://msdn.microsoft.com/library/0035be38-8a65-4693-bcc4-0a8dd7b1495b%28Office.15%29.aspx)|Especifica un tipo de entidad.|
|[EventType](eventtype-enumeration.md)|Especifica el tipo de evento que se ha generado.|
|[FileType](filetype-enumeration.md)|Especifica el formato en el que deben devolverse los documentos.|
|[GoToType](gototype-enumeration.md)|Especifica el tipo de lugar u objeto hacia el que se debe navegar.|
|[FilterType](filtertype-enumeration.md)|Especifica si se debe aplicar el filtrado desde la aplicación host al recuperar los datos.|
|[InitializationReason](initializationreason-enumeration.md)|Especifica si el complemento se acaba de insertar o si se encontraba en el documento con anterioridad.|
|[ItemType](http://msdn.microsoft.com/library/e0bb23fd-f360-4b0f-b72c-1cf08d4cab3f%28Office.15%29.aspx)|Especifica el tipo de un elemento.|
|[notificationMessageType](http://msdn.microsoft.com/library/ff00c89d-0019-4545-a95b-7ed0db712ce9%28Office.15%29.aspx)|Especifica el mensaje de notificación de una cita o un mensaje.|
|[ProjectProjectFields](projectprojectfields-enumeration.md)|Especifica los campos del proyecto que están disponibles como parámetro para el método [getProjectFieldAsync](projectdocument.getprojectfieldasync.md).|
|[ProjectResourceFields](projectresourcefields-enumeration.md)|Especifica los campos de recursos que están disponibles como parámetro para el método [getResourceFieldAsync](projectdocument.gettaskfieldasync.md).|
|[ProjectTaskFields](projecttaskfields-enumeration.md)|Especifica los campos de tarea que están disponibles como parámetro para el método [getTaskFieldAsync](projectdocument.gettaskfieldasync.md).|
|[ProjectViewTypes](projectviewtypes-enumeration.md)|Especifica los tipos de vistas que puede reconocer el método [getSelectedViewAsync](projectdocument.getselectedviewasync.md).|
|[RecipientType](http://msdn.microsoft.com/library/6e7c4029-6e52-47f6-98d2-4cd3ce7bd8b4%28Office.15%29.aspx)|Especifica el tipo de destinatario de una cita.|
|[ResponseType](http://msdn.microsoft.com/library/b3e723ca-4be0-4846-ad97-0eecab4355eb%28Office.15%29.aspx)|Especifica la respuesta para la invitación a una reunión.|
|[SelectionMode](selectionmode-enumeration.md)|Especifica si se va a seleccionar (resaltar) la ubicación a la que se va a dirigir (al usar el método [Document.goToByIdAsync](document.gotobyidasync.md)).|
|[SourceProperty](http://msdn.microsoft.com/library/6a209a7f-57cd-4dc3-869e-07b0f5928b28%28Office.15%29.aspx)|Especifica el origen de los datos devueltos por el método invocado.|
|[Tabla](table-enumeration.md)|Especifica los valores enumerados de la propiedad `cells:` en el parámetro _cellFormat_ de los [métodos de formato de tabla](../../docs/excel/format-tables-in-add-ins-for-excel.md).|
|[ValueFormat](valueformat-enumeration.md)|Especifica si se debe aplicar su formato correspondiente a los valores que devuelve el método que se ha invocado (por ejemplo, números y fechas).|

## Detalles de compatibilidad


La compatibilidad para cada enumeración difiere entre aplicaciones host de Office. Consulte la sección "Detalles de compatibilidad" del tema de cada enumeración para obtener información de compatibilidad de host.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Tipos de complementos**|Contenido, panel de tareas y Outlook|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|
