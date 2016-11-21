
# <a name="documentactiveviewchanged-event"></a>Evento Document.ActiveViewChanged
Se produce cuando el usuario cambia la vista actual del documento.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint|
|**Introducido en**|1.1|

```
Office.EventType.ActiveViewChanged
```


## <a name="remarks"></a>Comentarios

Para agregar un controlador de eventos para el evento **ActiveViewChanged** de un documento, use el método [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) del objeto **Document**. El controlador de eventos recibirá un argumento de tipo [ActiveViewChangedEventArgs](../../reference/shared/document.activeviewchangedeventargs.md).


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para Mac**|**Office para iPad**|
|:-----|:-----|:-----|:-----|:-----|
|**PowerPoint**|v||v|v|

|||
|:-----|:-----|
|**Introducido en**|1.1|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|
