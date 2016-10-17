
# <a name="document.gotobyidasync-method"></a>Método Document.goToByIdAsync
Va al objeto o la ubicación que se haya especificado en el documento.

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint y Word|
|**Disponible en los conjuntos de requisitos**|No en un conjunto|
|**Agregado en**|1.1|

```js
Office.context.document.goToByIdAsync(id, goToType, [,options], callback);
```


## <a name="parameters"></a>Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _id_|**string** o **number**|El identificador del objeto o la ubicación a la que dirigirse. Obligatorio.||
| _goToType_|[GoToType](../../reference/shared/gototype-enumeration.md)|El tipo de ubicación a la que dirigirse. Obligatorio.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _selectionMode_|[SelectionMode](../../reference/shared/selectionmode-enumeration.md)|Especifica si la ubicación especificada por el parámetro _id_ está seleccionada (resaltada).|**En Excel:**<br/> **Office.SelectionMode.Selected** selecciona todo el contenido del enlace o del elemento con nombre. <br/>**Office.SelectionMode.None**: para los enlaces de texto, se selecciona la celda; para enlaces de matrices, los enlaces de tablas y los elementos con nombre, se selecciona la primera celda de datos (no la primera celda de la fila de encabezado de las tablas).<br/><br/> **En PowerPoint:**<br/> **Office.SelectionMode.Selected**: selecciona el título de la diapositiva o el primer cuadro de texto de la diapositiva.<br/> **Office.SelectionMode.None**: no se selecciona nada.<br/><br/> **En Word:**<br/> **Office.SelectionMode.Selected**: se selecciona todo el contenido del enlace. <br/>**Office.SelectionMode.None** para los enlaces de texto; mueve el cursor al principio del texto; para los enlaces de matriz y los enlaces de la tabla, selecciona la primera celda de datos (no la primera celda de la fila de encabezado de las tablas).|
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **goToByIdAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la información siguiente.



|**Propiedad**|**Usar para**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Devolver la vista actual.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **objeto** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## <a name="remarks"></a>Comentarios

PowerPoint no es compatible con el método **goToByIdAsync** en **Vista Patrón**.


## <a name="example"></a>Ejemplo

 **Ir a un enlace por id. (Word y Excel)**

En el ejemplo siguiente se muestra cómo hacerlo:


-  **Cree un enlace de tabla** usando el método [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) como un enlace de muestra.
    
-  **Especifique dicho enlace** como el enlace al que dirigirse.
    
-  **Pase una función de devolución de llamada anónima** que devuelva el estado de la operación al parámetro _callback_ del método **goToByIdAsync**.
    
-  **Mostrar el valor** en la página del complemento.
    



```js
function gotoBinding() {
    //Create a new table binding for the selected table.
    Office.context.document.bindings.addFromSelectionAsync("table",{ id: "MyTableBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
              showMessage("Action failed with error: " + asyncResult.error.message);
           }
           else {
              showMessage("Added new binding with type: " + asyncResult.value.type +" and id: " + asyncResult.value.id);
           }
    });

    //Go to binding by id.
    Office.context.document.goToByIdAsync("MyTableBinding", Office.GoToType.Binding, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **Ir a una tabla en una hoja de cálculo (Excel)**

En el ejemplo siguiente se muestra cómo hacerlo:


-  **Especifique el nombre de una tabla** como la tabla a la que dirigirse.
    
-  **Pase una función de devolución de llamada anónima** que devuelva el estado de la operación al parámetro _callback_ del método **goToByIdAsync**.
    
-  **Mostrar el valor** en la página del complemento.
    



```js
function goToTable() {
    Office.context.document.goToByIdAsync("Table1", Office.GoToType.NamedItem, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **Ir a la diapositiva seleccionada actualmente por id. (PowerPoint)**

En el ejemplo siguiente se muestra cómo hacerlo:


-  **Obtenga el id.** de las diapositivas seleccionadas actualmente usando el método [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md).
    
-  **Especifique el Id. devuelto** como la diapositiva a la que dirigirse.
    
-  **Pase una función de devolución de llamada anónima** que devuelva el estado de la operación al parámetro _callback_ del método **goToByIdAsync**.
    
-  **Mostrar el valor** del objeto JSON stringified devuelto por `asyncResult.value`, que contiene información sobre las diapositivas seleccionadas, en la página del complemento.
    



```js
var firstSlideId = 0;
function gotoSelectedSlide() {
    //Get currently selected slide's id
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
    //Go to slide by id.
    Office.context.document.goToByIdAsync(firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```



 **Ir a la diapositiva por índice (PowerPoint)**

En el ejemplo siguiente se muestra cómo hacerlo:


-  **Especifique el índice** de la primera o última diapositiva, o de la positiva anterior o siguiente a la que dirigirse.
    
-  **Pase una función de devolución de llamada anónima** que devuelva el estado de la operación al parámetro _callback_ del método **goToByIdAsync**.
    
-  **Mostrar el valor** en la página del complemento.
    



```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```




## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|No en un conjunto|
|**Nivel de permisos mínimo**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint Online.|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Agregado|
