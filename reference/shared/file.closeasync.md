
# <a name="file.closeasync-method"></a>Método File.closeAsync
Cierra el archivo de documento.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Archivo|
|**Agregado en**|1.1|

```js
File.closeAsync(callback);
```


## <a name="parameters"></a>Parámetros


_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **object**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo [AsyncResult](../../reference/shared/asyncresult.md). Opcional.
    

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **closeAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Devuelve siempre **undefined** porque no hay ningún objeto o dato que recuperar.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **objeto** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## <a name="remarks"></a>Comentarios

No se permite que haya más de dos documentos en la memoria; de lo contrario, la operación [Document.getFileAsync](../../reference/shared/document.getfileasync.md) fallará. Use el método **File.closeAsync** para cerrar el archivo cuando haya terminado de trabajar con él.


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|Archivo|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint y Word en Office para iPad.|
|1.0|Agregado|
