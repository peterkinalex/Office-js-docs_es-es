
# Propiedad DocumentActiveViewChangedEventArgs.activeView
Se obtiene un valor de enumeración **ActiveView** que identifica el estado de la vista activa del documento; por ejemplo, determina si el usuario puede editar el documento.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint|
|**Agregado en**|1.1|

```
var myView = eventArgsObj.activeView;
```


## Valor devuelto

La [ActiveView](../../reference/shared/activeview-enumeration.md) de la vista que generó el evento.


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|v|v|v|

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint en Office para iPad.|
|1.1|Agregado|
