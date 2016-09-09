
# Propiedad BindingSelectionChangedEventArgs.startColumn
Obtiene el índice de la primera columna de la selección (de base cero).

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Modificado por última vez en**|1.1|

```
var startCol = eventArgsObj.startColumn;
```


## Valor devuelto

El índice de base cero de la primera columna de la selección empezando desde la columna más a la izquierda del enlace.


## Comentarios

Si el usuario hace una selección no contigua, se devuelven las coordenadas de la última selección contigua que hay en el enlace. 

Para Word, esta propiedad solo funcionará para los enlaces de la "tabla" [BindingType](../../reference/shared/bindingtype-enumeration.md). Si la tabla es del tipo "matriz", se devuelve **null**. Además, la llamada fallará si la tabla contiene celdas combinadas, porque la estructura de la tabla tiene que ser uniforme para que esta propiedad funcione correctamente.


## Ejemplo

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


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
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
