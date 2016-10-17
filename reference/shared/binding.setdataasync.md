
# <a name="binding.setdataasync-method"></a>Método Binding.setDataAsync
Escribe datos en la sección enlazada del documento que representa el objeto de enlace que se ha especificado.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Disponible en los [conjuntos de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Modificado por última vez en TableBindings**|1.1|

```js
bindingObj.setDataAsync(data [, options] ,callback);
```


## <a name="parameters"></a>Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _data_|<table><tr><td><b>string</b></td><td>Solo en Excel, Excel Online, Word y Word Online</td></tr><tr><td><b>array</b> (matriz de matrices – "matrix")</td><td>Solo en Excel y Word</td></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp161002">
  <b>TableData</b></a></td><td>Solo en Access, Excel y Word</td></tr><tr><td><b>HTML</b></td><td>Solo en Word y Word Online</td></tr><tr><td><b>Office Open XML</b></td><td>Solo en Word</td></tr></table>|Los datos que se definirán en la selección actual. Requerido.|**Modificado en:** 1.1. La compatibilidad para los complementos de contenido para Access requiere el conjunto de requisitos **TableBinding** 1.1 o posterior.|
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|Especifica cómo convertir los datos que se están estableciendo. ||
| _columnas_|**matriz de cadenas**| Especifica los nombres de columna.|**Agregado en:** v1.1. Solo para los enlaces de tablas de los complementos de contenido para Access.|
| _filas_|**Office.TableRange.ThisRow**|Especifica la cadena predefinida "thisRow" para establecer los datos de la fila seleccionada actualmente. |**Agregado en:** v1.1. Solo para los enlaces de tablas de los complementos de contenido para Access.|
| _startColumn_|**number**|Especifica la columna de inicio de base cero para un subconjunto de los datos. |Solo para los enlaces de tabla o matriz. Si se omiten, los datos se establecen a partir de la primera columna.|
| _startRow_|**number**|Especifica la fila de inicio de base cero para un subconjunto de los datos en el enlace. |Solo para los enlaces de tabla o matriz. Si se omiten, los datos se establecen a partir de la primera fila.|
| _tableOptions_|**object**|Para la tabla insertada, una lista de pares clave-valor que especifican [opciones de formato de tabla](../../docs/excel/format-tables-in-add-ins-for-excel.md), como fila de encabezado, fila de total y filas con bandas. |**Agregado en:** v1.1. **Admitido en:** Excel.|
| _cellFormat_|**object**|Para la tabla insertada, una lista de pares clave-valor que especifican un rango de celdas, filas o columnas y el [formato de celda](../../docs/excel/format-tables-in-add-ins-for-excel.md) que se debe aplicar a dicho rango.|**Agregado en:** v1.1. **Admitido en:** Excel, Excel Online.|
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha pasado al método **setDataAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Devuelve siempre **undefined** porque no hay ningún objeto o dato que recuperar.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **objeto** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## <a name="remarks"></a>Comentarios

El valor que se ha pasado para _data_ contiene los datos que se escribirán en el enlace. El tipo de valor que se pasa determina qué se escribirá, tal como se describe en la tabla siguiente.



|**_Valor _data**|**Datos escritos**|
|:-----|:-----|
|Una **cadena**|Se escribirá texto sin formato o cualquier cosa que pueda convertirse en una **string**.|
|Una matriz de matrices ("matriz")|Se escribirán datos tabulares sin encabezados. Por ejemplo, para escribir datos en tres filas de dos columnas, se puede transferir una matriz como esta: ` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`. Para escribir una sola columna de tres filas, transfiera una matriz como esta: `[["R1C1"], ["R2C1"], ["R3C1"]]`|
|Un objeto [TableData](../../reference/shared/tabledata.md)|Se escribirá una tabla con encabezados.|
Estas acciones específicas de aplicaciones también se pueden realizar al escribir datos en un enlace.

 **Para Word**, el parámetro _data_ se escribe en el enlace del siguiente modo:



|**_Valor _data**|**Datos escritos**|
|:-----|:-----|
|Una **cadena**|Se escribe el texto que se ha especificado.|
|Una matriz de matrices ("matrix") o un objeto **TableData**|Se escribe una tabla de Word.|
|HTML|Se escribe el contenido HTML que se ha especificado.
 >**Importante** Si parte del contenido HTML que se escribe no es válido, Word no generará un error, sino que escribirá el máximo contenido HTML posible y omitirá los datos no válidos.

|
|Office Open XML ("Open XML")|Se escribe el contenido XML que se ha especificado.|  **Para Excel**, el parámetro _data_ se escribe en el enlace del siguiente modo:



|**_Valor _data**|**Datos escritos**|
|:-----|:-----|
|Una **cadena**|El texto especificado se inserta como el valor de la primera celda enlazada. También puede especificar una fórmula válida para agregarla a la celda enlazada. Por ejemplo, al establecer _data_ en `"=SUM(A1:A5)"`, se calculará el total de los valores en el rango especificado. Sin embargo, cuando se establece una fórmula en la celda enlazada, después de hacerlo no puede leer la fórmula (o cualquier fórmula existente) de la celda enlazada. Si llama al método [Binding.getDataAsync](../../reference/shared/binding.getdataasync.md) en la celda enlazada para leer sus datos, el método puede devolver solo los datos que se muestran en la celda (resultado de la fórmula).|
|Una matriz de matrices ("matrix") y la forma coincide exactamente con la forma del enlace que se ha especificado|Se escribe el conjunto de filas y columnas. También puede especificar una matriz de matrices que contenga fórmulas válidas para agregarlas a las celdas enlazadas. Por ejemplo, al establecer _data_ en `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]`, se agregarán estas dos fórmulas a un enlace que contiene dos celdas. Igual que cuando se establece una fórmula en una única celda enlazada, no podrá leer las fórmulas agregadas (o fórmulas existentes) del enlace con el método **Binding.getDataAsync**, porque solo devuelve los datos que se muestran en las celdas enlazadas.|
|Un objeto **TableData** y la forma de tabla coincide con la tabla enlazada.|Se escribe el conjunto especificado de filas o encabezados si no se van a sobrescribir otros datos de las celdas contiguas. **Nota:** si especifica fórmulas en el objeto **TableData** que pasa al parámetro _data_, podría no obtener los resultados que espera debido a la característica "columnas calculadas" de Excel, que automáticamente duplica las fórmulas dentro de una columna. Para solucionar esto cuando quiere escribir _data_ que contienen fórmulas a una tabla enlazada, pruebe a especificar los datos como una matriz de matrices (en lugar de un objeto **TableData**) y especifique _coercionType_ como **Microsoft.Office.Matrix** o "matriz".|
 **Comentarios adicionales para Excel Online**


- El número total de celdas en el valor pasado al parámetro _data_ no puede ser superior a 20 000 en una sola llamada a este método.
    
- El número de _grupos de formato_ pasado al parámetro _cellFormat_ no puede ser superior a 100. Un único grupo de formato consta de un conjunto de formato aplicado a un rango de celdas especificado. Por ejemplo, la siguiente llamada pasa dos grupos de formato a _cellFormat_.
    
```js
  Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});

```

En todos los casos restantes, se devolverá un error.

El método **setDataAsync** escribirá los datos en un subconjunto de un enlace de matriz o tabla si se especifican los parámetros opcionales _startRow_ y _startColumn_, y estos definen a su vez un rango válido.


## <a name="example"></a>Ejemplo




```js
function setBindingData() {
    Office.select("bindings#MyBinding").setDataAsync('Hello World!', function (asyncResult) { });
}
```

Especificar el parámetro _coercionType_ opcional le permite especificar el tipo de datos que quiere escribir en un enlace. Por ejemplo, en Word, si quiere escribir HTML en un enlace de texto, puede especificar el parámetro _coercionType_ como `"html"` como se muestra en el ejemplo siguiente, que usa etiquetas HTML `<b>` para que "Hello" aparezca en negrita.




```js
function writeHtmlData() {
    Office.select("bindings#myBinding").setDataAsync("<b>Hello</b> World!", {coercionType: "html"}, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

En este ejemplo, la llamada a **setDataAsync** remite el parámetro _data_ como una matriz de matrices (para crear una sola columna con tres filas) y especifica la estructura de datos con el parámetro _coercionType_ como `"matrix"`.




```js
function writeBoundDataMatrix() {
    Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],{ coercionType: "matrix" }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

En la función `writeBoundDataTable` de este ejemplo, la llamada a **setDataAsync** remite el parámetro _data_ como un objeto **TableData** (para escribir tres columnas y tres filas) y especifica la estructura de datos con el parámetro _coercionType_ como `"table"`. 

En la función `updateTableData`, la llamada a **setDataAsync** remite de nuevo el parámetro _data_ como un objeto **TableData**, pero en forma de una única columna con un nuevo encabezado y tres filas, para actualizar los valores de la última columna de la tabla creada con la función `writeBoundDataTable`. El parámetro opcional de base cero _startColumn_ se especifica como 2 para reemplazar los valores de la tercera columna de la tabla.




```js
function writeBoundDataTable() {
    // Create a TableData object.
    var myTable = new Office.TableData();
    myTable.headers = ['First Name', 'Last Name', 'Grade'];
    myTable.rows = [['Kim', 'Abercrombie', 'A'], ['Junmin','Hao', 'C'],['Toni','Poe','B']];

    // Set myTable in the binding.
    Office.select("bindings#myBinding").setDataAsync(myTable, { coercionType: "table" }, 
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Error: '+ asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}

// Replace last column with different data.
function updateTableData() {
     var newTable = new Office.TableData();
     newTable.headers = ["Gender"];
     newTable.rows = [["M"],["M"],["F"]];
     Office.select("bindings#myBinding").setDataAsync(newTable, { coercionType: "table", startColumn:2 }, 
         function (asyncResult) {
             if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                 write('Error: '+ asyncResult.error.message);
         } else {
            write('Bound data: ' + asyncResult.value);
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
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|MatrixBindings, TableBindings, TextBindings|
|**Nivel de permisos mínimo**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|<ul><li>En los complementos para Access, se ha agregado compatibilidad para escribir datos de tabla.</li><li>En los complementos para Excel, se ha agregado compatibilidad para <a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">establecer el formato al escribir datos en un enlace de tabla</a> con los parámetros opcionales <span class="parameter" sdata="paramReference">tableOptions</span> y <span class="parameter" sdata="paramReference">cellFormat</span>.</li></ul>|
|1,0|Agregado|

## <a name="see-also"></a>Vea también



#### <a name="other-resources"></a>Otros recursos


[Enlazar a regiones en un documento u hoja de cálculo](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
