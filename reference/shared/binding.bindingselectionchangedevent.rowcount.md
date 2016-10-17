
# <a name="bindingselectionchangedeventargs.rowcount-property"></a>Propiedad BindingSelectionChangedEventArgs.rowCount
Obtiene la cantidad de filas seleccionadas.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Modificado por última vez en**|1.1|

```
var rwCount = eventArgsObj.rowCount;
```


## <a name="return-value"></a>Valor devuelto

El número de filas que se han seleccionado. Si se selecciona una sola celda, devolverá 1.


## <a name="remarks"></a>Comentarios

Si el usuario selecciona elementos que no son contiguos, devolverá el recuento del último grupo de elementos contiguos que haya seleccionado dentro del enlace. 

En Word, esta propiedad solo funciona con enlaces "table" de tipo [BindingType](../../reference/shared/bindingtype-enumeration.md). Si el enlace es de tipo "matrix", devolverá **null**. También se producirá un error en la llamada si la tabla contiene celdas combinadas, porque la estructura de la tabla debe ser uniforme para que esta propiedad funcione correctamente.


## <a name="example"></a>Ejemplo

En el ejemplo siguiente se agrega un controlador de eventos para el evento [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) al enlace con un [id](../../reference/shared/binding.id.md) de `myTable`. Cuando el usuario cambia la selección, en el controlador se muestran las coordenadas de la primera celda de la selección y el número de filas y columnas seleccionadas.


```js
function addSelectionHandler() {
    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addHandlerAsync("bindingSelectionChanged", myHandler);
    });
}

// Display selection start coordinates and row/column count.
function myHandler(bArgs) {
    write("Selection start row/col: " + bArgs.startRow + "," + bArgs.startColumn);
    write("Selection row count: " + bArgs.rowCount);
    write("Selection col count: " + bArgs.columnCount);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Ahora puede agregar y quitar controladores de eventos para el evento **SelectionChanged** en los complementos de contenido para Access.|
|1.0|Agregado|
