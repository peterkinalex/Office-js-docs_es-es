
# <a name="bindingselectionchangedeventargs-object"></a>Objeto BindingSelectionChangedEventArgs
Proporciona información sobre el enlace que ha generado el evento [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md).

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Modificado por última vez en TableBinding**|1.1|

```
Office.EventType.BindingSelectionChanged
```


## <a name="members"></a>Miembros


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[binding](../../reference/shared/binding.bindingselectionchangedevent.binding.md)|Obtiene un objeto [Binding](../../reference/shared/binding.md) que representa el enlace que ha generado el evento **SelectionChanged**.|
|[columnCount](../../reference/shared/binding.bindingselectionchangedevent.columncount.md)|Obtiene la cantidad de columnas seleccionadas.|
|[rowCount](../../reference/shared/binding.bindingselectionchangedevent.rowcount.md)|Obtiene la cantidad de filas seleccionadas.|
|[startRow](../../reference/shared/binding.bindingselectionchangedevent.startrow.md)|Obtiene el índice de la primera fila de la selección (de base cero).|
|[startColumn](../../reference/shared/binding.bindingselectionchangedevent.startcolumn.md)|Obtiene el índice de la primera columna de la selección (de base cero).|
|[type](../../reference/shared/binding.bindingselectionchangedevent.type.md)|Obtiene un valor de enumeración [EventType](../../reference/shared/eventtype-enumeration.md) que identifica el tipo de evento que se generó.|

## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para el enlace de tabla en los complementos para Access.|
|1.0|Agregado|
