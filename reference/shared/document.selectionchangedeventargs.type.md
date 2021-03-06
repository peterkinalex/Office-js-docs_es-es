
# <a name="documentselectionchangedeventargs.type-property"></a>Propiedad DocumentSelectionChangedEventArgs.type
Obtiene un valor de la enumeración **EventType** que identifica el tipo de evento que se ha generado.

|||
|:-----|:-----|
|**Hosts:**|Excel y Word|
|**Agregado en**|1.1|

```
var myEvent = eventArgsObj.type;
```


## <a name="return-value"></a>Valor devuelto

La enumeración [EventType](../../reference/shared/eventtype-enumeration.md) del evento que se ha generado.


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
