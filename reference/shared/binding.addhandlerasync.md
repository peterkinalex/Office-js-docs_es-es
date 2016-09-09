
# Método Binding.addHandlerAsync
Agrega un controlador al enlace para el tipo de evento especificado.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|BindingEvents|
|**Modificado por última vez en**|1.1|

```
bindingObj.addHandlerAsync(eventType, handler [, options], callback);
```


## Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|Especifica el tipo de evento que se debe agregar. Necesario. Para un evento de objetos **Binding**, el parámetro _eventType_ se puede especificar como **Office.EventType.BindingSelectionChanged**, **Office.EventType.BindingDataChanged** o los valores de texto correspondientes de estas enumeraciones.||
| _handler_|**object**|La función del controlador de eventos que se debe agregar.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **addHandlerAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Siempre devuelve **undefined** porque no hay ningún objeto ni datos que recuperar al agregar un controlador de eventos.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **object** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## Comentarios

Puede agregar varios controladores de eventos para el _eventType_ especificado siempre que cada controlador de eventos tenga un nombre exclusivo.


## Ejemplo

El ejemplo de código siguiente llama al método [select](../../reference/shared/office.select.md) del objeto **Office** para obtener acceso al enlace con el identificador "MyBinding" y, a continuación, llama al método **addHandlerAsync** para agregar una función de controlador para el evento [bindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md) de dicho enlace.


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




## Detalles de compatibilidad


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
|**Disponible en los conjuntos de requisitos **|BindingEvents|
|**Nivel de permisos mínimo**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para los complementos para Access.|
|1.0|Agregado|
