
# <a name="binding.getdataasync-method"></a>Método Binding.getDataAsync
Devuelve los datos que contiene el enlace.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Disponible en los [conjuntos de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Modificado por última vez en TableBindings**|1.1|

```
bindingObj.getDataAsync([, options] , callback );
```


## <a name="parameters"></a>Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|Especifica cómo convertir los datos que se están estableciendo. ||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|Especifica si se debe aplicar su formato correspondiente a los valores que se devuelven (por ejemplo, números y fechas).||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|Especifica si se debe aplicar un filtro al recuperar los datos.||
| _filas_|**Office.TableRange.ThisRow**| Especifica la cadena predefinida "thisRow" para obtener los datos de la fila seleccionada actualmente.|Solo para los enlaces de tablas de los complementos de contenido para Access.|
| _startRow_|**number**|Para los enlaces de matriz o tabla, especifica la fila de inicio de base cero para un subconjunto de los datos del enlace. ||
| _startColumn_|**number**|Para los enlaces de matriz o tabla, especifica la columna de inicio de base cero para un subconjunto de los datos del enlace. ||
| _rowCount_|**number**|Para los enlaces de matriz o tabla, especifica el número de filas que se deben desplazar desde _startRow_. ||
| _columnCount_|**number**|Para los enlaces de matriz o tabla, especifica el número de columnas que se deben desplazar desde _startColumn_.||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **Binding.getDataAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Tener acceso a los valores del enlace especificado. Si se determina el parámetro _coercionType_ (y la llamada se completa correctamente), los datos se devolverán en el formato que se describe en el tema de la enumeración [CoercionType](../../reference/shared/coerciontype-enumeration.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **objeto** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## <a name="remarks"></a>Comentarios

Si se omite un parámetro opcional, se usará el siguiente valor predeterminado (cuando resulte aplicable al tipo y al formato de los datos).



|**Parámetro**|**Predeterminado**|
|:-----|:-----|
| _coercionType_|El tipo original del enlace, sin coerción.|
| _valueFormat_|Datos sin formato.|
| _filterType_|Todos los valores (sin filtrado).|
| _startRow_|La primera fila.|
| _startColumn_|La primera columna.|
| _rowCount_|Todas las filas.|
| _columnCount_|Todas las columnas.|
Si se realiza la llamada desde [MatrixBinding](../../reference/shared/binding.matrixbinding.md) o [TableBinding](../../reference/shared/binding.tablebinding.md), y se especifican los parámetros opcionales **startRow**, _startColumn_, _rowCount_ y _columnCount_ (siempre que estos definan un rango válido de elementos contiguos), el método _getDataAsync_ devolverá un subconjunto de los valores enlazados.


## <a name="example"></a>Ejemplo




```
function showBindingData() {
    Office.select("bindings#MyBinding").getDataAsync(function (asyncResult) {
        write(asyncResult.value)
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



Existe una importante diferencia de comportamiento entre el uso de `"table"` y `"matrix"`_coercionType_ con el método **Binding.getDataAsync**, con respecto a los datos formateados con filas de encabezado, tal como se muestra en los dos ejemplos siguientes. Estos ejemplos de código muestran funciones de controlador de eventos para el evento [Binding.SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md).

Si especifica la `"table"` _coercionType_, la propiedad [TableData.rows](../../reference/shared/tabledata.rows.md) (`result.value.rows` en el siguiente ejemplo de código) devuelve una matriz que solo contiene las filas de cuerpo de la tabla. Por lo que su fila 0 será la primera fila que no es de encabezado de la tabla.




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'table', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value.rows[0][0]); 
            } 
            else 
                write(result.error.message); 
    }); 
}     
// Function that writes to a div with id='message' on the page. 
function write(message){ 
    document.getElementById('message').innerText += message; 
}
```

Sin embargo, si especifica la `"matrix"` _coercionType_, `result.value` en el siguiente ejemplo de código devuelve una matriz que contiene el encabezado de tabla en la fila 0. Si el encabezado de tabla contiene varias filas, entonces todas ellas se incluyen en la matriz `result.value` como filas separadas antes de que se incluyan las filas del cuerpo de la tabla.




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'matrix', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value[1][0]); 
            } 
            else 
                write(result.error.message); 
    }); 
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
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|MatrixBindings, TableBindings, TextBindings|
|**Nivel de permisos mínimo**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para los enlaces de tabla en los complementos para Access.|
|1.0|Agregado|

## <a name="see-also"></a>Vea también



#### <a name="other-resources"></a>Otros recursos


[Enlazar a regiones en un documento u hoja de cálculo](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
