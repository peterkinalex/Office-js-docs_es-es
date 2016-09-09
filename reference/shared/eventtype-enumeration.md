
# Enumeración EventType
Especifica el tipo de evento que se ha generado y se devuelve desde la propiedad **type** de un objeto _EventName_**EventArgs**.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project y Word|
|**Modificado por última vez en Selección**|1.1|

```js
Office.EventType
```


## Miembros


**Valores**


|Enumeración|Valor|Descripción|
|:-----|:-----|:-----|
|Office.EventType.ActiveViewChanged|"documentActiveViewChanged"|Se ha generado un evento [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md).|
|Office.EventType.DocumentSelectionChanged|"documentSelectionChanged"|Se ha generado un evento [Document.SelectionChanged](../../reference/shared/document.selectionchanged.event.md).|
|Office.EventType.BindingSelectionChanged|"bindingSelectionChanged"|Se ha generado un evento [Binding.BindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md).|
|Office.EventType.BindingDataChanged|"bindingDataChanged"|Se ha generado un evento [Binding.BindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md).|
|Office.EventType.DataNodeDeleted|"nodeDeleted"|Se generó un evento [CustomXmlPart.dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md).|
|Office.EventType.DataNodeInserted|"nodeInserted"|Se generó un evento [CustomXmlPart.dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md).|
|Office.EventType.DataNodeReplaced|"nodeReplaced"|Se generó un evento [CustomXmlPart.dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md).|
|Office.EventType.SettingsChanged|"settingsChanged"|Se ha generado un evento [Settings.settingsChanged](../../reference/shared/settings.settingschangedevent.md).|

## Observaciones


 >**Nota**:  Los complementos para Project admiten los tipos de evento  **Office.EventType.ResourceSelectionChanged**,  **Office.EventType.TaskSelectionChanged** y  **Office.EventType.ViewSelectionChanged**.


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta enumeración es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**PowerPoint**|v|v||
|**Project**|v|||
|**Word**|v||v|

|||
|:-----|:-----|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



|**Versión**|**Cambios**|
|:-----|:-----|
|1.1| Enumeración de Added Office.EventType.ActiveViewChanged para el nuevo evento **Document.ActiveViewChanged**.|
|1.0|Agregado|
