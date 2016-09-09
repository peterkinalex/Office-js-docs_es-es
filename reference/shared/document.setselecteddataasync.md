
# Método Document.setSelectedDataAsync
Escribe datos en la selección actual del documento.

|||
|:-----|:-----|
|**Hosts:** Access, Excel, PowerPoint, Project, Word y Word Online|**Tipos de complementos: ** Panel de tareas y contenido|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selección|
|**Modificado por última vez en**|1.1|

```js
Office.context.document.setSelectedDataAsync(data [, options], callback(asyncResult));
```


## Parámetros

|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _data_|Los datos pueden tener uno de los siguientes tipos de datos:<ul><li><b>string</b> (Office.CoercionType.Text): solo se aplica a Excel, Excel Online, PowerPoint, PowerPoint Online, Word y Word Online.</li><li><b>array</b> de matrices (Office.CoercionType.Matrix): solo se aplica a Excel, Word y Word Online.</li><li>[TableData](../../reference/shared/tabledata.md) (Office.CoercionType.Table): solo para Access, Excel, Word y Word Online</li><li><b>HTML</b> (Office.CoercionType.Matrix): solo se aplica a Word y Word Online.</li><li><b>Office Open XML</b> (Office.CoercionType.Matrix): solo se aplica a Word y Word Online.</li><li><b>Base64 encoded image stream</b>  (Office.CoercionType.Image): solo se aplica a Excel, PowerPoint, Word y Word Online.</li></ul>|Los datos que se definirán en la selección actual. Necesario.|**Modificado en:** 1.1. La compatibilidad para los complementos de contenido para Access requiere el conjunto de requisitos **Selection** 1.1 o posterior. La compatibilidad para establecer datos de imagen requiere el conjunto de requisitos **ImageCoercion** 1.1 o posterior. Para establecerlo para la activación de la aplicación, use:<br/><br/>`<Requirements>`<br/>&nbsp;&nbsp;`<Sets DefaultMinVersion="1.1">`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`<Set Name="ImageCoercion"/>`<br/>&nbsp;&nbsp;`</Sets>`<br/>`</Requirements>`<br/><br/>La detección en tiempo de ejecución de la capacidad ImageCoercion se puede llevar a cabo con el siguiente código:<br/><br/>`if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1')) {)) {`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`// insertViaImageCoercion();`<br/>`} else {`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`// insertViaOoxml();`<br/>`}`|
| _options_|**object**|Especifica un conjunto de [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). El objeto de opciones puede contener las siguientes propiedades para establecer las opciones:<br/><ul><li>coercionType (<b><a href="735eaab6-5e31-4bc2-add5-9d378900a31b.htm">CoercionType</a></b>): especifica cómo convertir los datos que se están estableciendo. Si no se establece esta opción, se usa el valor predeterminado coercionType de Office.CoercionType.Text.</li><li>tableOptions (<b>object</b>): para la tabla insertada, una lista de pares clave-valor que especifican <a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">opciones de formato de tabla</a>, como fila de encabezado, fila de total y filas con bandas. </li><li>cellFormat (<b>object</b>): para la tabla insertada, una lista de pares clave-valor que especifican un rango de celdas, filas o columnas y el <a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">formato de celda</a> que se debe aplicar a dicho rango. </li><li>imageLeft (<b>number</b>): esta opción es aplicable para la inserción de imágenes. Indica la ubicación de inserción en relación con el lado izquierdo de la diapositiva de PowerPoint y su relación con la celda seleccionada actualmente en Excel. Este valor se ignora para Word. Este valor está en puntos.</li><li>imageTop (<b>number</b>): esta opción es aplicable para la inserción de imágenes. Indica la ubicación de inserción en relación con la parte superior de la diapositiva de PowerPoint y su relación con la celda seleccionada actualmente en Excel. Este valor se ignora para Word. Este valor está en puntos.</li><li>imageWidth (<b>number</b>): esta opción es aplicable para la inserción de imágenes. Indica el ancho de la imagen. Si esta opción se proporciona sin el valor imageHeight, se escalará la imagen para que coincida con el valor del ancho de imagen. Si se proporciona el ancho y la altura de la imagen, la imagen se ajustará de forma correspondiente. Si no se proporciona ni la altura ni el ancho de la imagen, se usará el tamaño de imagen y la relación de aspecto predeterminados. Este valor está en puntos.</li><li>imageHeight (<b>number</b>): esta opción es aplicable para la inserción de imágenes. Indica el alto de la imagen. Si esta opción se proporciona sin el valor imageWidth, se escalará la imagen para que coincida con el valor de la altura de la imagen. Si se proporciona el ancho y la altura de la imagen, la imagen se ajustará de forma correspondiente. Si no se proporciona ni la altura ni el ancho de la imagen, se usará el tamaño de imagen y la relación de aspecto predeterminados. Este valor está en puntos.</li><li>asyncContext (<b>object \| value</b>): un objeto definido por el usuario que está disponible en la propiedad asyncContext del objeto <a href="540c114f-0398-425c-baf3-7363f2f6bc47.htm">AsyncResult</a>. Use esta opción para proporcionar un objeto o valor para el <b>AsyncResult</b> si la devolución de llamada es una función con nombre.</li></ul>|Las opciones _tableOptions_ y _cellFormat_ se han agregado en la versión 1.1 y son compatibles con Excel 2013 y Excel Online.<br/><br/>Las opciones _imageLeft_ y _ImageTop_ son compatibles con Excel y PowerPoint.|
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha pasado al método **setSelectedDataAsync**, la propiedad [AsyncResult.value](../../reference/shared/asyncresult.value.md) siempre devuelve **undefined** porque no hay ningún objeto o dato que recuperar.


## Observaciones

El valor transferido para _data_ contiene los datos que se deben escribir en la selección actual. Si el valor es:


-  **Una string** Se insertará texto sin formato o cualquier cosa que pueda convertirse en una **string**.
    
    
    
    En Excel, también puede especificar _data_ como una fórmula válida para agregar esa fórmula a la celda seleccionada. Por ejemplo, establecer _data_ a `"=SUM(A1:A5)"` totalizará los valores en el rango especificado. En cambio, cuando se establece una fórmula en la celda dependiente, después de hacerlo, no se puede leer desde la celda dependiente la fórmula agregada (o cualquier fórmula preexistente). Si se llama al método [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) en la celda seleccionada para leer sus datos, el método puede devolver solo los datos que se muestran en la celda (el resultado de la fórmula).
    
-  **Una matriz de matrices ("matriz"):** Se insertarán datos tabulares sin encabezados. Por ejemplo, para escribir datos en tres filas de dos columnas, puede pasar una matriz como esta: `[["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`. Para escribir una sola columna de tres filas, pase una matriz como esta: `[["R1C1"], ["R2C1"], ["R3C1"]]`
    
    
    
    En Excel, también puede especificar _data_ como una matriz de matrices que contiene fórmulas válidas para agregarlas a las celdas seleccionadas. Por ejemplo, si no se sobrescribirán otros datos, establecer _data_ a `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]` agregará esas dos fórmulas a la selección. Igual que cuando se establece una fórmula en una sola celda como "texto", no se pueden leer las fórmulas agregadas (o las fórmulas existentes) después de que se han configurado, solo se pueden leer los resultados de las fórmulas.
    
-  **Un objeto [TableData](../../reference/shared/tabledata.md):** Se insertará una tabla con encabezados.
    
    
    
     **Nota:** en Excel, si especifica fórmulas en el objeto **TableData** que pasa al parámetro _data_, podría no obtener los resultados que espera debido a la característica "columnas calculadas" de Excel, que automáticamente duplica las fórmulas dentro de una columna. Para solucionar esto cuando quiere escribir _data_ que contienen fórmulas a una tabla seleccionada, pruebe a especificar los datos como una matriz de matrices (en lugar de un objeto **TableData**), y especifique _coercionType_ como **Microsoft.Office.Matrix** o "matriz".
    
 **Comportamientos específicos de la aplicación**

Además, al escribir datos en una selección, se aplican las siguientes acciones específicas de la aplicación.

 **Para Word**


- Si no hay ninguna selección y el punto de inserción es una ubicación válida, se inserta el elemento _data_ en el punto de inserción tal y como se describe a continuación:
    
      - If  _data_ is a string, the specified text is inserted.
    
  - Si _data_ es una matriz de matrices o un objeto **TableData**, se inserta una tabla de Word nueva.
    
  - Si _data_ es HTML, se inserta el HTML especificado.
    
     >**Importante**:  Si algún HTML que inserta no es válido, Word no mostrará un error. Word insertará tanto código HTML como pueda y omitirá los datos no válidos.
  - Si _data_ es Office Open XML, se inserta el XML especificado.
    
  - Si _data_ es una secuencia de imágenes codificada con base64, se inserta la imagen especificada.
    
- Si hay una selección, se sustituirá por el elemento _data_ especificado de acuerdo con las mismas reglas anteriores.
    
-  **Insertar imágenes**: Las imágenes insertadas se colocan alineadas. Los parámetros **imageLeft** e **imageTop** se ignoran. La relación de aspecto de la imagen siempre está bloqueada. Si solo se indica uno de los parámetros **imageWidth** e **imageHeight**, el otro valor se escalará automáticamente para conservar la relación de aspecto original.
    
 **Para Excel**


- Si se selecciona solo una celda:
    
      - If  _data_ is a string, the specified text is inserted as the value of the current cell.
    
  - Si _data_ es una matriz de matrices ("matrix"), se inserta el conjunto de filas y columnas especificado, si no se sustituyen otros datos de las celdas adyacentes.
    
  - Si _data_ es un objeto **TableData**, se inserta una tabla de Excel nueva con el conjunto de filas y encabezados especificado, si no se sustituyen otros datos de las celdas adyacentes.
    
- Si se seleccionan varias celdas y la forma no coincide con la de _data_, se devuelve un error.
    
- Si se seleccionan varias celdas y la forma de la selección coincide exactamente con la de _data_, se actualizan los valores de las celdas seleccionadas según los valores de _data_.
    
-  **Insertar imágenes**: Las imágenes insertadas son flotantes. Los parámetros de posición **imageLeft** e **imageTop** dependen de las celdas seleccionadas actualmente. Los valores negativos **imageLeft** e **imageTop** están permitidos y es posible que Excel los reajuste para colocar la imagen dentro de una hoja de cálculo. La relación de aspecto de la imagen está bloqueada a menos que se indiquen los parámetros **imageWidth** e **imageHeight**. Si solo se indica uno de los parámetros **imageWidth** e **imageHeight**, el otro valor se escalará automáticamente para conservar la relación de aspecto original.
    
En todos los casos restantes, se devolverá un error.

 **Para Excel Online**

Además de los comportamientos descritos anteriormente para Excel, al escribir los datos en Excel Online se aplican los límites siguientes. 


- El número total de celdas que puede escribir en una hoja de cálculo con el parámetro _data_ no puede exceder las 20.000 en una sola llamada a este método.
    
- El número de _grupos de formato_ pasado al parámetro _cellFormat_ no puede ser superior a 100. Un único grupo de formato consta de un conjunto de formato aplicado a un rango de celdas especificado. Por ejemplo, la siguiente llamada pasa dos grupos de formato a _cellFormat_.
    

```js
  Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```

 **Para PowerPoint**

Las imágenes insertadas son flotantes. Los parámetros de posición **imageLeft** e **imageTop** son opcionales; no obstante, si se indican, ambos deben estar presentes. Si solo se proporciona un valor, se ignorará. Los valores negativos **imageLeft** e **imageTop** están permitidos y pueden colocar una imagen fuera de una diapositiva. Si no se indica ningún parámetro opcional y la diapositiva tiene un marcador de posición, la imagen reemplazará el marcador de la diapositiva. La relación de aspecto de la imagen se bloqueará a menos que se indiquen los parámetros **imageWidth** e **imageHeight**. Si solo se indica uno de los parámetros **imageWidth** e **imageHeight**, el otro valor se escalará automáticamente para conservar la relación de aspecto original.


## Ejemplo

En el ejemplo siguiente se define el texto o la celda seleccionada en "Hello World!" y, si falla, se muestra el valor de la propiedad [error.message](../../reference/shared/error.message.md).


```js
function writeText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                 write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



Especificar el parámetro _coercionType_ opcional le permite especificar el tipo de datos que quiere escribir en una selección. En el ejemplo siguiente se escriben datos como una matriz de tres filas y dos columnas, especificando _coercionType_ como `"matrix"` para esa estructura de datos y, si eso falla, se muestra el valor de la propiedad [error.message](../../reference/shared/error.message.md).




```js
function writeMatrix() {
    Office.context.document.setSelectedDataAsync([["Red", "Rojo"], ["Green", "Verde"], ["Blue", "Azul"]], {coercionType: Office.CoercionType.Matrix}
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



En el ejemplo siguiente se escriben datos como una tabla de una columna con un encabezado y cuatro filas, especificando _coercionType_ como `"table"` para esa estructura de datos y, si eso falla, se muestra el valor de la propiedad [error.message](../../reference/shared/error.message.md).




```js
function writeTable() {
    // Build table.
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [['Berlin'], ['Roma'], ['Tokyo'], ['Seattle']];

    // Write table.
    Office.context.document.setSelectedDataAsync(myTable, {coercionType: Office.CoercionType.Table},
        function (result) {
            var error = result.error
            if (result.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



 En Word, si quiere escribir HTML en la selección, puede especificar el parámetro _coercionType_ como `"html"` tal y como se muestra en el ejemplo siguiente, que usa etiquetas HTML `<b>` para que "Hello" aparezca en negrita.




```js
function writeHtmlData() {
    Office.context.document.setSelectedDataAsync("<b>Hello</b> World!", {coercionType: Office.CoercionType.Html}, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

En Word, PowerPoint y Excel, si quiere escribir una imagen en la selección, puede especificar el parámetro _coercionType_ como `"image"`, tal y como se muestra en el ejemplo siguiente. Tenga en cuenta que Word ignora imageLeft e imageTop.




```js
function insertPictureAtSelection(base64EncodedImageStr) {

    Office.context.document.setSelectedDataAsync(base64EncodedImageStr, {
       coercionType: Office.CoercionType.Image,
       imageLeft: 50,
       imageTop: 50,
       imageWidth: 100,
       imageHeight: 100
       },
       function (asyncResult) {
           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
               console.log("Action failed with error: " + asyncResult.error.message);
           }
       });
}
```


## Detalles de compatibilidad


Una marca de verificación (![Símbolo de verificación](../../images/mod_off15_checkmark.png)) en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**

||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|![Símbolo de verificación](../../images/mod_off15_checkmark.png)|||
|**Excel**|![Símbolo de verificación](../../images/mod_off15_checkmark.png)|![Símbolo de verificación](../../images/mod_off15_checkmark.png)|![Símbolo de verificación](../../images/mod_off15_checkmark.png)|
|**PowerPoint**|![Símbolo de verificación](../../images/mod_off15_checkmark.png)|![Símbolo de verificación](../../images/mod_off15_checkmark.png)|![Símbolo de verificación](../../images/mod_off15_checkmark.png)|
|**Word**|![Símbolo de verificación](../../images/mod_off15_checkmark.png)|![Símbolo de verificación](../../images/mod_off15_checkmark.png)|![Símbolo de verificación](../../images/mod_off15_checkmark.png)|


|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|Selección|
|**Nivel de permisos mínimo**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|En Word y Word Online se ha agregado compatibilidad para escribir los datos como una secuencia de imágenes codificada en base64.|
|1.1|En Word Online se ha agregado compatibilidad para escribir _data_ como una **array** de matrices (matriz) y **TableData** (tabla).|
|1.1|En Excel, PowerPoint y Word en Office para iPad se ha agregado el mismo nivel de compatibilidad que para Excel, PowerPoint y Word en el escritorio de Windows.|
|1.1|En Word Online, se ha agregado compatibilidad para escribir _data_ como una **string** (texto).|
|1.1|Agregada compatibilidad con la [definición de formato al insertar tablas](../../docs/excel/format-tables-in-add-ins-for-excel.md) con complementos para Excel, a través de los parámetros opcionales _tableOptions_ y _cellFormat_.|
|1.1|Se ha agregado compatibilidad para la escritura de datos de tabla en complementos para Access.|
|1.0|Agregado|
