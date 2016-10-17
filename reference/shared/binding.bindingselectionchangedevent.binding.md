
# <a name="bindingselectionchangedeventargs.binding-property"></a>Propiedad BindingSelectionChangedEventArgs.binding
Obtiene un objeto **Binding** que representa el enlace que generó el evento **SelectionChanged**.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Modificado por última vez en**|1.1|

```
var myBinding = eventArgsObj.binding;
```


## <a name="return-value"></a>Valor devuelto

Un objeto [Binding](../../reference/shared/binding.md) que representa el enlace que ha generado el evento [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md).


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v|v|v|

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
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Ahora puede agregar y quitar controladores de eventos para el evento **SelectionChanged** en los complementos de contenido para Access.|
|1.0|Agregado|
