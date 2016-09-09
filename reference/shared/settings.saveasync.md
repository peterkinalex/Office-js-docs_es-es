
# Método Settings.saveAsync
Mantiene la copia en memoria del contenedor de propiedades de configuración en el documento.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Configuración|
|**Modificado por última vez en**|1.1|

```js
Office.context.document.settings.saveAsync(callback);
```


## Parámetros



_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **objeto**

&nbsp;&nbsp;&nbsp;&nbsp;Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**. Opcional.

    



## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **saveAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Devuelve siempre **undefined** porque no hay ningún objeto o dato que recuperar.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **object** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## Comentarios

Al inicializar un complemento, se cargarán todas las configuraciones que haya guardado. Esto significa que, durante la sesión, solo podrá usar los métodos [set](../../reference/shared/settings.set.md) y [get](../../reference/shared/settings.get.md) para trabajar con la copia en memoria del contenedor de propiedades de configuración. Si desea guardar la configuración para que esté disponible la próxima vez que use el complemento, use el método **saveAsync**.


 >**Nota**: el método **saveAsync** conserva el contenedor de propiedades de configuración en memoria en el archivo del documento; no obstante, los cambios que se realicen en el propio documento se guardan solo si el usuario (o la configuración **Autorrecuperación**) guarda el documento en el sistema de archivos.

El método [refreshAsync](../../reference/shared/settings.refreshasync.md) solo es útil en escenarios de coautoría (que solo son compatibles con Word) cuando otras instancias del mismo complemento pueden cambiar la configuración y dichos cambios se deben poner a disposición de todas las instancias.


## Ejemplo




```js
function persistSettings() {
    Office.context.document.settings.saveAsync(function (asyncResult) {
        write('Settings saved with status: ' + asyncResult.status);
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
