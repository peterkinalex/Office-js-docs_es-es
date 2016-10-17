
# <a name="binding.bindingdatachanged-event"></a>Evento Binding.bindingDataChanged
Se produce al cambiar los datos en el enlace.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Modificado por última vez en BindingEvents**|1.1|

```js
Office.EventType.BindingDataChanged
```


## <a name="remarks"></a>Comentarios

Para agregar un controlador de eventos para el evento **BindingDataChanged** de un enlace, use el método [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) del objeto **Binding**. El controlador de eventos recibirá un argumento de tipo [BindingDataChangedEventArgs](../../reference/shared/binding.bindingdatachangedeventargs.md).


## <a name="example"></a>Ejemplo




```js
function addEventHandlerToBinding() {
    Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
}

function onBindingDataChanged(eventArgs) {
    write("Data has changed in binding: " + eventArgs.binding.id);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|BindingEvents|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad

|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para este evento en los complementos para Access.|
|1.0|Agregado|
