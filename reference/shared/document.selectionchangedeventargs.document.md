
# <a name="documentselectionchangedeventargs.document-property"></a>Propiedad DocumentSelectionChangedEventArgs.document
Obtiene un objeto **Document** que representa el documento que generó el evento **SelectionChanged**.

|||
|:-----|:-----|
|**Hosts:**|Excel y Word|
|**Agregado en**|1.1|




```js
var myDoc = eventArgsObj.document;
```


## <a name="return-value"></a>Valor devuelto

Un objeto [Document](../../reference/shared/document.md) que representa el documento que ha generado el evento [SelectionChanged](../../reference/shared/document.selectionchanged.event.md).


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.0|Agregado|
