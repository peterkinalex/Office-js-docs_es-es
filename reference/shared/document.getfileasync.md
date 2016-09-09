
# Método Document.getFileAsync
Devuelve el archivo de documento entero en segmentos de hasta 4194304 bytes (4 MB). En cuanto a los complementos para iOS, el segmento de archivo puede tener un tamaño máximo de 65536 bytes (64 KB). Tenga en cuenta que, si especifica un tamaño del segmento del archivo superior al límite permitido, se producirá el error "Error interno". 

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Archivo|
|**Modificado por última vez en Archivo**|1.1|

```js
Office.context.document.getFileAsync(fileType [, options], callback);
```


## Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _fileType_|[FileType](../../reference/shared/filetype-enumeration.md)|Especifica el formato en el que se devolverá el archivo. Necesario.<br/><table><tr><th>Host</th><th>Tipo de archivo admitido</th></tr><tr><td>Excel Online</td><td>Office.FileType.Compressed</td></tr><tr><td>PowerPoint para el escritorio de Windows</td><td>Office.FileType.Compressed, Office.FileType.Pdf</td></tr><tr><td>Word para el escritorio de Windows, MAC y iPad</td><td>Office.FileType.Compressed, Office.FileType.Pdf, Office.FileType.Text</td></tr><tr><td>Word Online</td><td>Office.FileType.Compressed, Office.FileType.Pdf, Office.FileType.Text</td></tr><tr><td>PowerPoint Online</td><td>Office.FileType.Compressed, Office.FileType.Pdf</td></tr></table>|**Modificado en** 1.1; consulte [Historial de compatibilidad](#historial-de-compatibilidad)|
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _sliceSize_|**number**|Especifica el tamaño del segmento deseado (en bytes) hasta 4194304 bytes (4MB). Si no se especifica, se usará un tamaño predeterminado del segmento de 4194304 bytes (4MB). ||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **getFileAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Tener acceso al objeto [File](../../reference/shared/file.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **object** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## Comentarios

En el caso de los complementos que se ejecutan en aplicaciones host de Office que no son de Office para iOS, el método **getFileAsync** admite la obtención de archivos en segmentos de hasta 4194304 bytes (4 MB). En el caso de los complementos que se ejecutan en aplicaciones de Office para iOS, el método **getFileAsync** admite la obtención de archivos en segmentos de hasta 65536 (64 KB).

El parámetro _fileType_ puede especificarse con las enumeraciones o los valores de texto siguientes.


**Enumeración FileType**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.FileType.Compressed|"compressed"|Devuelve el documento completo (.docx, .pptx o .xslx) en el formato Office Open XML (OOXML) como una matriz de bytes.|
|Office.FileType.Pdf|"pdf"|Devuelve todo el documento en formato PDF como matriz de bytes.|
|Office.FileType.Text|"text"|Devuelve solo el texto del documento como un **string**. |
No se permite que haya más de dos documentos en la memoria; de lo contrario, la operación **getFileAsync** fallará. Use el método [File.closeAsync](../../reference/shared/file.closeasync.md) para cerrar el archivo cuando haya terminado de trabajar con él.


## Ejemplo: obtener un documento en formato Office Open XML ("comprimido")

En el siguiente ejemplo se obtiene el documento en formato Office Open XML ("comprimido") en segmentos de 65536 bytes (64 KB). Nota: la implementación de `app.showNotification` en este ejemplo procede de la plantilla de Visual Studio para los complementos de Office.


```js
function getDocumentAsCompressed() {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ }, 
        function (result) {
            if (result.status == "succeeded") {
            // If the getFileAsync call succeeded, then
            // result.value will return a valid File Object.
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);

            // Get the file slices.
            getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
            else {
            app.showNotification("Error:", result.error.message);
            }
    });
}

function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
    file.getSliceAsync(nextSlice, function (sliceResult) {
        if (sliceResult.status == "succeeded") {
            if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                return;
            }

            // Got one slice, store it in a temporary array.
            // (Or you can do something else, such as
            // send it to a third-party server.)
            docdataSlices[sliceResult.value.index] = sliceResult.value.data;
            if (++slicesReceived == sliceCount) {
               // All slices have been received.
               file.closeAsync();
               onGotAllSlices(docdataSlices);
            }
            else {
                getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
        }
            else {
                gotAllSlices = false;
                file.closeAsync();
                app.showNotification("getSliceAsync Error:", sliceResult.error.message);
            }
    });
}

function onGotAllSlices(docdataSlices) {
    var docdata = [];
    for (var i = 0; i < docdataSlices.length; i++) {
        docdata = docdata.concat(docdataSlices[i]);
    }

    var fileContent = new String();
    for (var j = 0; j < docdata.length; j++) {
        fileContent += String.fromCharCode(docdata[j]);
    }

    // Now all the file content is stored in 'fileContent' variable,
    // you can do something with it, such as print, fax...
}

```


## Ejemplo: obtener un documento en formato PDF

En el ejemplo siguiente se obtiene el documento en formato PDF.


```js
Office.context.document.getFileAsync(Office.FileType.Pdf,
    function(result) {
        if (result.status == "succeeded") {
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);
            // Now, you can call getSliceAsync to download the files, as described in the previous code segment (compressed format).
            
            myFile.closeAsync();
        }
        else {
            app.showNotification("Error:", result.error.message);
        }
}
);


```


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||v||
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|Archivo|
|**Nivel de permisos mínimo**|[ReadAllDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1| En PowerPoint Online se ha agregado compatibilidad para **Office.FileType.Pdf** como el parámetro _fileType_.|
|1.1| En PowerPoint Online se ha agregado compatibilidad para **Office.FileType.Compressed** como el parámetro _fileType_.|
|1.1| En Word Online se ha agregado compatibilidad para **Office.FileType.Text** como el parámetro _fileType_.|
|1.1| En Excel Online se ha agregado compatibilidad para **Office.FileType.Compressed** como el parámetro _fileType_.|
|1.1| En Word Online se ha agregado compatibilidad para **Office.FileType.Compressed** y **Office.FileType.Pdf** como el parámetro _fileType_.|
|1.1|En PowerPoint y Word en Office para iPad, se ha agregado compatibilidad para todos los valores **FileType** como el parámetro _fileType_.|
|1.1|En Word y PowerPoint para el escritorio de Windows, se ha agregado compatibilidad para **Office.FileType.Pdf** como el parámetro _fileType_.|
|1.0|Agregado|
