
# <a name="customxmlpart.datanodedeleted-event"></a>Evento CustomXmlPart.dataNodeDeleted
Se produce cuando se elimina un nodo.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Modificado por última vez en**|1.1|

```
Office.EventType.DataNodeDeleted
```


## <a name="remarks"></a>Comentarios

Para agregar un controlador de eventos para el evento **dataNodeDeleted**, use el método [addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md) del objeto **CustomXmlPart**.


## <a name="example"></a>Ejemplo




```js
function addNodeDeletedEvent() {
    Office.context.document.customXmlParts.getByIdAsync("{3BC85265-09D6-4205-B665-8EB239A8B9A1}", function (result) {
        var xmlPart = result.value;
        xmlPart.addHandlerAsync(Office.EventType.DataNodeDeleted, function (eventArgs) {
            write("A node has been deleted.");
        });
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```




## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este evento es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este evento.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||||
|**Excel**||||
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|CustomXmlParts|
|**Nivel de permisos mínimo**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word en Office para iPad.|
|1.0|Agregado|
