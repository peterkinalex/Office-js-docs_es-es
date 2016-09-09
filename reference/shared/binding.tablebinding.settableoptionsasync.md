
# Método TableBinding.setTableOptionsAsync
Actualiza las opciones de formato de tabla en la tabla enlazada.

|||
|:-----|:-----|
|**Hosts:**|Excel|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|No en un conjunto|
|**Agregado en**|1.1|

```
bindingObj.setTableOptionsAsync(tableOptions [,options] , callback);
```


## Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _tableOptions_|**object**|Literal de objeto que contiene una lista de pares nombre-valor de propiedad que define las opciones de tabla que se aplicarán. Obligatorio.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **goToByIdAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la información siguiente.



|**Propiedad**|**Usar para...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Siempre devuelve **undefined** porque no hay ningún objeto ni datos que recuperar al establecer las opciones de tabla.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **object** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## Ejemplo

En el ejemplo siguiente se muestra cómo hacerlo:


-  **Crear un literal de objeto** que especifique las [opciones de formato de tabla](../../docs/excel/format-tables-in-add-ins-for-excel.md) que actualizar en la tabla enlazada.
    
-  **Llamar a setTableOptions** en una tabla enlazada previamente (que tenga un **id** de `myBinding`) que pase el objeto con la opción de formato como el parámetro _tableOptions_.
    

```js
function updateTableFormatting(){
    var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

    Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
}
```




## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|No en un conjunto.|
|**Nivel de permisos mínimo**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel en Office para iPad.|
|1.1|Agregado|
