

# Método Settings.refreshAsync
Lee toda la configuración que se conserva en el documento y actualiza la copia de dicha configuración del complemento de contenido o del panel de tareas, que se conserva en la memoria.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Configuración|
|**Modificado por última vez en**|1.1|

```js
Office.context.document.settings.refreshAsync(callback);
```


## Parámetros

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **objeto**

&nbsp;&nbsp;&nbsp;&nbsp;Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.

    



## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **refreshAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Tener acceso a un objeto [Settings](../../reference/shared/settings.md) con los valores actualizados.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **object** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## Comentarios

Este método resulta útil en escenarios de coautoría de Word y PowerPoint cuando varias instancias del mismo complemento trabajan en un mismo documento. Como cada complemento trabaja en una copia en memoria de la configuración cargada desde el documento cuando un usuario lo abre, es posible que los valores de configuración de los distintos usuarios no estén sincronizados. Esto puede ocurrir siempre que una instancia del complemento llame al método [Settings.saveAsync](../../reference/shared/settings.saveasync.md) para guardar la configuración completa del usuario en el documento. Si desea actualizar los valores de configuración para todos los usuarios, llame al método **refreshAsync** desde el controlador de eventos para el evento [settingsChanged](../../reference/shared/settings.settingschangedevent.md) del complemento.

También puede llamar al método **refreshAsync** desde los complementos creados para Excel, aunque no tiene mucho sentido hacerlo, dado que estos programas no admiten la coautoría.


## Ejemplo




```js
function refreshSettings() {
    Office.context.document.settings.refreshAsync(function (asyncResult) {
        write('Settings refreshed with status: ' + asyncResult.status);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|Configuración|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint Online.|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para las configuraciones personalizadas en complementos de contenido para Access.|
|1.0|Agregado|
