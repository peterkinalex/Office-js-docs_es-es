
# Propiedad AsyncResult.value
Obtiene la carga o el contenido de una operación asincrónica, si la hay.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project y Word|
|**Modificado por última vez en**|1.1|

```js
var dataValue = asyncResult.value;
```


## Valor devuelto

Devuelve el valor de la solicitud en el momento en que se realiza la llamada asincrónica. 


 >**Nota**:  el resultado que devuelve la propiedad **value** para un método "Async" específico varía en función de la finalidad y el contexto de dicho método. Para determinar el resultado que la propiedad **value** debe devolver para un método "Async", consulte la sección "Valor de devolución de llamada" del tema correspondiente al método. Para obtener un listado completo de los métodos "Async", consulte la sección "Notas" del tema del objeto [AsyncResult](../../reference/shared/asyncresult.md).


## Observaciones

Si lo desea, puede obtener acceso al objeto **AsyncResult** de la función que se ha remitido como argumento al parámetro _callback_ de un método "Async", por ejemplo, los métodos [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) y [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) del objeto **Document**.


## Ejemplo




```js
function getData() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Table, function(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
        }
        else {
            write(asyncResult.value);
        }
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

||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|**OWA para dispositivos**|**Office para Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||v||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Project**|v|||||
|**Word**|v|v|v|||

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas y Outlook|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint Online.|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para los complementos para Access.|
|1.0|Agregado|
