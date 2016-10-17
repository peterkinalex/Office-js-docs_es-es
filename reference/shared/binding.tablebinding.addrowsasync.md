
# <a name="tablebinding.addrowsasync-method"></a>Método TableBinding.addRowsAsync
Agrega filas y valores a una tabla.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Modificado por última vez en**|1.1|

```js
bindingObj.addRowsAsync(rows, [,options], callback);
```


## <a name="parameters"></a>Parámetros

_rows_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **Array**

&nbsp;&nbsp;&nbsp;&nbsp;Una matriz de matrices que contiene una o varias filas de datos para agregar a la tabla. Obligatorio.
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **object**

&nbsp;&nbsp;&nbsp;&nbsp;Especifica los [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) siguientes.
    
&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **array, boolean, null, number, object, string o undefined**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto [AsyncResult](../../reference/shared/asyncresult.md) sin sufrir modificaciones. Opcional.<br/><br/>

_callback_<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **object**
    
&nbsp;&nbsp;&nbsp;&nbsp;Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo [AsyncResult](../../reference/shared/asyncresult.md). Opcional.



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _rows_|**array**|Una matriz de matrices que contiene una o varias filas de datos para agregar a la tabla. Obligatorio.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **addRowsAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la información siguiente.



|**Propiedad**|**Usar para**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Devuelve siempre **undefined** porque no hay ningún objeto o dato que recuperar.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **objeto** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## <a name="remarks"></a>Comentarios

El resultado de éxito o error de una operación de **addRowsAsync** es atómico, es decir, toda la operación de agregar filas debe completarse correctamente. De lo contrario, se revertirá totalmente (y la propiedad **AsyncResult.status** que se ha devuelto con la devolución de llamada informará del error):


- Cada fila de la matriz que remita como argumento _data_ deberá tener el mismo número de columnas que la tabla que se está actualizando. De no ser así, fallará toda la operación.
    
- Todas las filas y las celdas de la matriz deben agregarse correctamente a las filas y celdas correspondientes de la tabla, en la o las filas que se acaban de crear. Si no se establece correctamente alguna de estas filas o celdas por cualquier motivo, toda la operación fallará.
    
 **Comentarios adicionales para Excel Online**

El número total de celdas en el valor pasado al parámetro _rows_ no puede ser superior a 20 000 en una sola llamada a este método.


## <a name="example"></a>Ejemplo




```js
function addRowsToTable() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        var binding = asyncResult.value;
        binding.addRowsAsync([["6", "k"], ["7", "j"]]);
    });
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
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|TableBindings|
|**Nivel de permisos mínimo**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para la escritura de datos de tabla en complementos para Access.|
|1.0|Agregado|
