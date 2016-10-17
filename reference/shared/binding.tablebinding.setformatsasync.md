
# <a name="tablebinding.setformatsasync-method"></a>Método TableBinding.setFormatsAsync
Establece o actualiza el formato de los elementos y datos especificados en la tabla enlazada.

|||
|:-----|:-----|
|**Hosts:**|Excel|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|No en un conjunto|
|**Agregado en**|1.1|

```
bindingObj.setFormatsAsync(cellFormat [,options] , callback);
```


## <a name="parameters"></a>Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _cellFormat_|**array**|Una matriz que contiene uno o varios objetos de JavaScript que especifican las celdas que deben considerarse celdas de destino y el formato que se les aplicará. Obligatorio.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **goToByIdAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la información siguiente.



|**Propiedad**|**Usar para**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Siempre devuelve **undefined** porque no hay ningún objeto ni datos que recuperar al establecer los formatos.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **objeto** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## <a name="remarks"></a>Comentarios

 **Especificar el parámetro cellFormat**

Use el parámetro _cellFormat_ para establecer o cambiar los valores de formato de celda (como el ancho, el alto, la fuente, el fondo, la alineación, etc.). El valor que pase como parámetro _cellFormat_ deberá ser una **array** que contenga una lista de uno o varios objetos JavaScript que especifiquen las celdas que deben considerarse celdas de destino (`cells:`) y los formatos que se les aplicará (`format:`).

Cada objeto JavaScript en la matriz _cellFormat_ tiene la siguiente forma:

 `{cells:{`_cell_range_`}, format:{`_format_definition_`}}`

La propiedad `cells:` especifica el rango al que quiere aplicar formato mediante uno de los siguientes valores:


**Rangos admitidos en la propiedad cells**


|**configuración del rango de cells**|**Descripción**|
|:-----|:-----|
| `{row: i}`|Especifica el rango que se extiende hasta la fila ith de datos de la tabla.|
| `{column: i}`|Especifica el rango que se extiende hasta la columna ith de datos de la tabla.|
| `{row: i, column: j}`|Especifica el rango de celdas desde la fila ith hasta la columna jth de datos de la tabla.|
| `Office.Table.All`|Especifica toda la tabla, incluidos los encabezados de columna, los datos y los totales (si resulta aplicable).|
| `Office.Table.Data`|Especifica solo los datos de la tabla (no los encabezados ni los totales).|
| `Office.Table.Headers`|Especifica solo la fila de encabezado.|


La propiedad `format:` especifica los valores que corresponden a un subconjunto de las opciones de configuración disponibles en el cuadro de diálogo **Formato de celdas** de Excel (haga clic con el botón derecho > **Formato de celdas** o **Inicio** > **Formato** > **Formato de celdas**).

Especifique el valor de la propiedad `format:` como una lista de uno o más pares _nombre_ - _valor_ de propiedad en un literal de objeto de JavaScript. El _nombre de propiedad_ especifica el nombre de la propiedad de formato que se va a establecer, y _valor_ especifica el valor de dicha propiedad. Se pueden especificar varios valores para un formato concreto, como el color y el tamaño de una fuente. A continuación, puede consultar tres ejemplos de valores de la propiedad `format:`:




```
//Set cells: font color to green and size to 15 points.
format: {fontColor : "green", fontSize : 15}
```




```
//Set cells: border to dotted blue.
format: {borderStyle: "dotted", borderColor: "blue"}
```




```
//Set cells: background to red and alignment to centered.
format: {backgroundColor: "red", alignHorizontal: "center"}
```

Puede especificar formatos de número introduciendo la cadena "código" del formato de número en la propiedad `numberFormat:`. Las cadenas de formato numérico que puede especificar se corresponden con las que se pueden establecer en Excel mediante la categoría **Personalizada** en la pestaña **Número** del cuadro de diálogo **Formato de celdas**. El ejemplo siguiente muestra cómo dar formato a un número como porcentaje con dos decimales:




```
format: {numberFormat:"0.00%"}
```

Para obtener más información, consulte cómo [crear un formato de número personalizado](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1#BM1).



 **Especificar un destino único**

El siguiente ejemplo muestra un valor _cellFormat_ que establece el color de fuente de la fila de encabezado en rojo.




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: Office.Table.Headers, format: {fontColor: "red"}}], 
    function (asyncResult){});
```

 **Especificar varios destinos**

El método **setFormatsAsync** permite dar formato a varios destinos dentro de la tabla enlazada en una única llamada de función. Para ello, debe pasar una lista de objetos en la matriz _cellFormat_ para cada destino al que desea dar formato. Por ejemplo, la siguiente línea de código establecerá en amarillo el color de fuente de la primera fila y un borde blanco y texto en negrita para la cuarta celda de la tercera fila.




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

Para establecer el formato de las tablas al escribir datos, use los parámetros opcionales _tableOptions_ y _cellFormat_ de los métodos [Document.setSelectedDataAsync](http://msdn.microsoft.com/library/4c1e13e9-b61a-47df-836c-3ca9aba4ca1c%28Office.15%29.aspx) o [TableBinding.setDataAsync](http://msdn.microsoft.com/library/5b6ecf6f-c57f-4c0d-9605-59daee8fde13%28Office.15%29.aspx).

La configuración de formatos con los parámetros opcionales de los métodos **Document.setSelectedDataAsync** y **TableBinding.setDataAsync** solo es válida para establecer el formato al escribir datos por primera vez. Si desea realizar cambios de formato después de escribir los datos, use los métodos siguientes:


- Para actualizar el formato de las celdas, como el color y el estilo de la fuente, use el método **TableBinding.setFormatsAsync** (este método).
    
- Para actualizar las opciones de la tabla, como las filas con bandas y los botones de filtro, use el método [TableBinding.setTableOptions](../../reference/shared/binding.tablebinding.settableoptionsasync.md).
    
- Para borrar el formato, use el método [TableBinding.clearFormats](../../reference/shared/binding.tablebinding.clearformatsasync.md).
    
 **Comentarios adicionales para Excel Online**

El número de _grupos de formato_ pasado al parámetro _cellFormat_ no puede ser superior a 100. Un único grupo de formato consta de un conjunto de formato aplicado a un rango de celdas especificado. Por ejemplo, la siguiente llamada pasa dos grupos de formato a _cellFormat_.




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});

```

Para obtener más detalles y ejemplos, consulte [Format tables in add-ins for Excel](../../docs/excel/format-tables-in-add-ins-for-excel.md) (Dar formato a las tablas de los complementos para Excel).


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**||**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|v||v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|No en un conjunto.|
|**Nivel de permisos mínimo**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel en Office para iPad.|
|1.1|Agregado|
