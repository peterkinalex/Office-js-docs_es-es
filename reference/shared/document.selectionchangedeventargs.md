
# <a name="documentselectionchangedeventargs-object"></a>Objeto DocumentSelectionChangedEventArgs
Proporciona información sobre el documento que generó el evento [SelectionChanged](../../reference/shared/document.selectionchanged.event.md).

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint y Word|
|**Agregado en**|1.1|

```

```


## <a name="members"></a>Miembros


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[document](../../reference/shared/document.selectionchangedeventargs.document.md)|Obtiene un objeto **Document** que representa el documento que generó el evento **SelectionChanged**.|
|[type](../../reference/shared/document.selectionchangedeventargs.type.md)|Obtiene un valor de la enumeración **EventType** que identifica el tipo de evento que se ha generado.|

## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.0|Agregado|
