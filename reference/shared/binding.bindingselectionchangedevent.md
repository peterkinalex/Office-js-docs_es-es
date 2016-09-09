
# Evento Binding.bindingSelectionChanged
Se genera al cambiar la selección en el enlace.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|BindingEvents|
|**Modificado por última vez en Selección**|1.1|

```
Office.EventType.BindingSelectionChanged
```

## Observaciones

Para agregar un controlador de eventos para el evento **BindingSelectionChanged** de un enlace, use el método [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) del objeto **Binding**. El controlador de eventos recibirá un argumento de tipo [BindingSelectionChangedEventArgs](../../reference/shared/binding.bindingselectionchangedeventargs.md).


## Ejemplo




```
function addEventHandlerToBinding() {
 Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
}

function onBindingSelectionChanged(eventArgs) {
    write(eventArgs.binding.id + " has been selected.");
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este evento es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este evento.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|BindingEvents|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad





****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para este evento en los complementos para Access.|
|1.0|Agregado|
