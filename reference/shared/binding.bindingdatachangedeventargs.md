
# <a name="bindingdatachangedeventargs-object"></a>Objeto BindingDataChangedEventArgs
Proporciona información sobre el enlace que generó el evento [DataChanged](../../reference/shared/binding.bindingdatachangedevent.md).

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Modificado por última vez en BindingEvents**|1.1|

```js
Office.EventType.BindingDataChanged
```


## <a name="members"></a>Miembros


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[binding](../../reference/shared/binding.bindingdatachangedeventargs.binding.md)|Obtiene un objeto [Binding](../../reference/shared/binding.md) que representa el enlace que generó el evento **DataChanged**.|
|[type](../../reference/shared/binding.bindingdatachangedeventargs.type.md)|Obtiene un valor de enumeración [EventType](../../reference/shared/eventtype-enumeration.md) que identifica el tipo de evento que se generó.|

## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para este evento en los complementos para Access.|
|1.0|Agregado|
