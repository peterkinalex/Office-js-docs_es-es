
# Método TableBinding.addColumnsAsync
Agrega columnas y valores a una tabla.

|||
|:-----|:-----|
|**Hosts:**|Excel y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Modificado por última vez en**|1,0|

```
bindingObj.addColumnsAsync(data [, options], callback);
```


## Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _data_|**array** o [TableData](../../reference/shared/tabledata.md)|Una matriz de matrices ("matriz") o un objeto **TableData** que contiene una o varias filas de datos para agregarlas a la tabla. Obligatorio.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **addColumnsAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Devuelve siempre **undefined** porque no hay ningún objeto o dato que recuperar.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **object** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## Comentarios

Para agregar una o varias columnas a partir de los valores de los datos y los encabezados especificados, envíe un objeto **TableData** como parámetro _data_. Para agregar una o varias columnas a partir únicamente de los datos especificados, envíe una matriz de matrices ("matriz") como parámetro _data_.

El correcto funcionamiento o el fallo de una acción **addColumnAsync** es atómico. Es decir, toda la acción de adición de columnas tiene que ser correcta o se deshará la acción completamente (y la propiedad **AsyncResult.status** que se devuelve a la devolución de llamada informará de un fallo):


- Cada fila de la matriz que se transmite como el argumento _data_ debe tener la misma cantidad de filas que la tabla que se va a actualizar. Si no, fallará toda la acción.
    
- Todas las filas y las celdas de la matriz tienen que agregar correctamente esa fila o celda a la tabla en las columnas recién agregadas. Si, por cualquier motivo, no se define alguna de las filas o las celdas, fallará toda la acción.
    
- Si se transmite un objeto **TableData** como argumento de datos, la cantidad de filas de encabezados debe coincidir con la de la tabla que se va a actualizar.
    
**Comentarios adicionales para Excel Online**

El número total de celdas en el objeto **TableData** pasado al parámetro _data_ no puede ser superior a 20.000 en una sola llamada a este método.


## Ejemplo

En el ejemplo siguiente, se agrega una sola columna con tres filas a una tabla enlazada con [id](../../reference/shared/binding.id.md)`"myTable"` mediante la transmisión de un objeto **TableData** como el argumento _data_ del método **addColumnsAsync**. Para que funcione correctamente, la tabla que se va a actualizar debe tener tres filas.


```js
// Add a column to a binding of type table by passing a TableData object.
function addColumns() {
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```

En el ejemplo siguiente, se agrega una sola columna con tres filas a una tabla enlazada con [id](../../reference/shared/binding.id.md)`myTable` mediante la transmisión de una matriz de matrices como el argumento _data_ del método **addColumnsAsync**. Para que funcione correctamente, la tabla que se va a actualizar debe tener tres filas.




```js
// Add a column to a binding of type table by passing an array of arrays.
function addColumns() {
    var myTable = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|TableBindings|
|**Nivel de permisos mínimo**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.0|Agregado|
