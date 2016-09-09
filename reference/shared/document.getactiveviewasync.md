
# Método Document.getActiveViewAsync
 Devuelve el estado de la vista actual de la presentación (edición o lectura).

|||
|:-----|:-----|
|**Hosts:** Excel, PowerPoint y Word|**Tipos de complementos:** Panel de tareas, contenido|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|ActiveView|
|**Agregado en ActiveView**|1.1|

```
Office.context.document.getActiveViewAsync([,options], callback);
```


## Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **getActiveViewAsync**, la propiedad [AsyncResult.value](../../reference/shared/asyncresult.value.md) devuelve el estado de la vista actual de la presentación. El valor devuelto puede ser `edit` o `read`. `edit` corresponde a cualquiera de las vistas en las que se pueden editar diapositivas, como **Normal** o **Vista Esquema**. `read` corresponde a **Presentación con diapositivas** o **Vista de lectura**.


## Comentarios

Puede desencadenar un evento al cambiar la vista.


## Ejemplo

Para obtener la vista de la presentación actual, debe escribir una función de devolución de llamada que devuelva ese valor. En el ejemplo siguiente se muestra cómo hacerlo:


-  **Pase una función de devolución de llamada anónima** que devuelva el tipo de vista al parámetro _callback_ del método **getActiveViewAsync**.
    
-  **Mostrar el valor** en la página del complemento.
    

```js
function getFileView() {
    // Get whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage(asyncResult.value);
        }
    });
}
```




## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|||v|
|**PowerPoint**|v|v|v|
|**Word**|||v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|ActiveView|
|**Agregado en ActiveView**|1.1|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad





****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Agregado.|
