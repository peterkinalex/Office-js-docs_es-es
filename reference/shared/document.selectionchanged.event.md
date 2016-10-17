
# <a name="document.selectionchanged-event"></a>Evento Document.SelectionChanged
Se produce al cambiar la selección en el documento.

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint y Word|
|**Introducido en**|1.1|

```
Office.EventType.DocumentSelectionChanged
```

## <a name="remarks"></a>Comentarios

Para agregar un controlador de eventos para el evento **SelectionChanged** de un documento, use el método [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) del objeto **Document**.


## <a name="example"></a>Ejemplo




```
function addEventHandlerToDocument() {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, MyHandler);
}

function MyHandler(eventArgs) {
    doSomethingWithDocument(eventArgs.document);
}

```




## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
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
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.0|Agregado|
