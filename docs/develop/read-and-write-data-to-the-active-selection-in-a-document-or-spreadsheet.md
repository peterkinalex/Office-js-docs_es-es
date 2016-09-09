
# Leer y escribir datos en la selección activa de un documento u hoja de cálculo

El objeto [Document](../../reference/shared/document.md) expone métodos que permiten leer y escribir en la selección actual del usuario en documentos u hojas de cálculo. Para ello, el objeto **Document** proporciona los métodos **getSelectedDataAsync** y **setSelectedDataAsync**. En este tema también se describe cómo leer, escribir y crear controladores de eventos para detectar los cambios realizados en la selección del usuario.

El método **getSelectedDataAsync** solo funciona en la selección actual del usuario. Si necesita guardar la selección en el documento, para que la misma selección esté disponible para leer y escribir en las distintas sesiones de ejecución del complemento, tiene que agregar un enlace con el método [Bindings.addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155.aspx) (o crear un enlace con uno de los otros métodos "addFrom" del objeto [Bindings](http://msdn.microsoft.com/en-us/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1.aspx)). Para obtener información sobre la creación de un enlace a una región de un documento y sobre cómo leer y escribir en él, vea [Enlazar a regiones de un documento u hoja de cálculo](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


### Lectura de los datos seleccionados


En el ejemplo siguiente se muestra cómo obtener datos de una selección en un documento con el método [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md).


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

En este ejemplo, el primer parámetro _coercionType_ se especifica como **Office.CoercionType.Text** (el parámetro también se puede especificar con la cadena literal `"text"`). Esto quiere decir que la propiedad [value](../../reference/shared/asyncresult.status.md) del objeto [AsyncResult](../../reference/shared/asyncresult.md) disponible del parámetro _asyncResult_ de la función de devolución de llamada devolverá una cadena (**string**) que contendrá el texto seleccionado en el documento. Especificar tipos distintos de coerción dará como resultado valores diferentes. [Office.CoercionType](../../reference/shared/coerciontype-enumeration.md) es una enumeración de los valores de tipos de coerción disponibles. **Office.CoercionType.Text** da como resultado la cadena "text".


 >**Sugerencia**   **¿Cuándo se debe usar la tabla coercionType en comparación con la matriz para el acceso a los datos?** Si necesita que los datos tabulares seleccionados se amplíen de forma dinámica al agregar filas y columnas y debe trabajar con encabezados de tabla, debe usar el tipo de datos de tabla (especificando el parámetro _coercionType_ del método **getSelectedDataAsync** como `"table"` o **Office.CoercionType.Table**). La adición de filas y columnas en la estructura de datos se admite tanto en datos de matriz como de tabla, pero la anexión de filas y columnas solo se admite para los datos de tabla. Si no tiene previsto agregar filas y columnas y los datos no requieren la funcionalidad de encabezados, debe usar el tipo de datos de matriz (especificando el parámetro  _coercionType_ del método **getSelecteDataAsync** como `"matrix"` o **Office.CoercionType.Matrix**), que ofrece un modelo más sencillo de interacción con los datos.

La función anónima transferida a la función como segundo parámetro _callback_ se ejecuta cuando se completa la operación **getSelectedDataAsync**. Se llama a la función con un único parámetro (_asyncResult_) que contiene el resultado y el estado de la llamada. Si la llamada produce errores, la propiedad [error](../../reference/shared/asyncresult.context.md) del objeto **AsyncResult** proporciona acceso al objeto [Error](../../reference/shared/error.md). Puede comprobar el valor de las propiedades [Error.name](../../reference/shared/error.name.md) y [Error.message](../../reference/shared/error.message.md) para determinar el motivo por el que la acción establecida produjo errores. En caso contrario, se mostrará el texto seleccionado en el documento.

La propiedad [AsyncResult.status](../../reference/shared/asyncresult.error.md) se usa en la instrucción **if** para comprobar si la llamada se realizó correctamente o no. [Office.AsyncResultStatus](../../reference/shared/asyncresultstatus-enumeration.md) es una enumeración de valores de propiedad **AsyncResult.status** disponibles. **Office.AsyncResultStatus.Failed** da como resultado la cadena "failed" (y también se puede especificar como esa cadena literal).


### Escritura de datos en la selección


En el ejemplo siguiente se muestra cómo configurar la selección para que muestre "Hola a todos!".


```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Pasar diferentes tipos de objetos para el parámetro  _data_ obtendrá resultados distintos. El resultado depende de qué esté seleccionado actualmente en el documento, de la aplicación que hospede su complemento y de si los datos pasados se pueden forzar a la selección actual.

La función anónima pasada al método [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) como parámetro _callback_ se ejecuta cuando se completa la llamada asincrónica. Cuando se escriben los datos en la selección con el método **setSelectedDataAsync**, el parámetro _asyncResult_ de la devolución de llamada solo proporciona acceso al estado de la llamada y al objeto [Error](../../reference/shared/error.md) si la llamada produce errores.

 **Nota:** A partir de Excel 2013 SP1 y la compilación correspondiente de Excel Online, ahora puede [establecer el formato al escribir una tabla en la selección actual](../../docs/excel/format-tables-in-add-ins-for-excel.md).


### Detección de cambios en la selección


En el ejemplo siguiente se muestra cómo detectar cambios en la selección cuando se agrega un controlador de eventos con el método [Document.addHandlerAsync](../../reference/shared/document.addhandlerasync.md) para el evento [SelectionChanged](../../reference/shared/document.selectionchanged.event.md) en el documento.


```
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){} 
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

El primer parámetro  _eventType_ especifica el nombre del evento al que se debe suscribir. Transferir la cadena `"documentSelectionChanged"` de este parámetro equivale a transferir el tipo de evento **Office.EventType.DocumentSelectionChanged** de la enumeración [Office.EventType](../../reference/shared/eventtype-enumeration.md).

La función `myHander()` que se pasa a la función como segundo parámetro _handler_ es un controlador de eventos que se ejecuta cuando se cambia la selección en el documento. Se llama a la función con un parámetro único, _eventArgs_, que incluye una referencia a un objeto [DocumentSelectionChangedEventArgs](../../reference/shared/document.selectionchangedeventargs.md) cuando se completa la operación asincrónica. Puede usar la propiedad [DocumentSelectionChangedEventArgs.document](../../reference/shared/document.selectionchangedeventargs.document.md) para obtener acceso al documento que generó el evento.


 >**Nota**  Puede agregar varios controladores de eventos a un evento determinado llamando de nuevo al método  **addHandlerAsync** y pasando una función de controlador de eventos adicional para el parámetro _handler_. Esto funcionará correctamente siempre que el nombre de cada función del controlador de eventos sea único.


### Desactivación de la detección de cambios en la selección


En el ejemplo siguiente se muestra cómo dejar de escuchar el evento [Document.SelectionChanged](../../reference/shared/document.selectionchanged.event.md) a través de una llamada al método [document.removeHandlerAsync](../../reference/shared/document.removehandlerasync.md).


```
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

El nombre de función  `myHandler` que se transfiere a la función como segundo parámetro _handler_ especifica el controlador de eventos que se quitará del evento **SelectionChanged**.


 >**Importante:** Si el parámetro opcional _handler_ se omite cuando se llama al método **removeHandlerAsync**, se quitarán todos los controladores de eventos del _eventType_ especificado.

