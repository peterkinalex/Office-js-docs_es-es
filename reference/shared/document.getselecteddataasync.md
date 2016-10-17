
# <a name="document.getselecteddataasync-method"></a>Método Document.getSelectedDataAsync
Lee los datos incluidos en la selección actual del documento.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project y Word|
|**Disponible en los conjuntos de requisitos**|Selección|
|**Modificado por última vez en Selección**|1.1|

```js
Office.context.document.getSelectedDataAsync(coercionType [, options], callback); 
```


## <a name="parameters"></a>Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)<br/><table><tr><td></td><td><b>Soporte del host</b></td></tr><tr><td><b>Office.CoercionType.Text</b> (cadena)</td><td>Solo en Excel, Excel Online, PowerPoint, PowerPoint Online, Word y Word Online</td></tr><tr><td><b>Office.CoercionType.Matrix</b> (matriz de matrices)</td><td>Solo en Excel, Word y Word Online</td></tr><tr><td><b>Office.CoercionType.Table</b> (objeto [TableData](../../reference/shared/tabledata.md))</td><td>Solo en Access, Excel, Word y Word Online</td></tr><tr><td><b>Office.CoercionType.Html</b></td><td>Solo en Word.</td></tr><tr><td><b>Office.CoercionType.Ooxml</b> (Office Open XML)</td><td>Solo en Word y Word Online</td></tr><tr><td><b>Office.CoercionType.SlideRange</b></td><td>Solo en PowerPoint y PowerPoint Online</td></tr></table>|El tipo de la estructura de datos que se debe devolver. Obligatorio.||
| _options_|**object**<br/><table><tr><td><i>valueFormat</i></td><td><b>[ValueFormat](../../reference/shared/valueformat-enumeration.md)</b></td><td>Especifica si se devuelve el resultado con sus valores de número o de fecha con o sin formato.</td><td></td></tr><tr><td><i>filterType</i></td><td>[FilterType](../../reference/shared/filtertype-enumeration.md)</td><td>Especifica si se aplica el filtrado al recuperar los datos. Opcional.</td><td>Este parámetro se ignora en los documentos de Word.</td></tr><tr><td><i>asyncContext</i></td><td><b>array</b>, <b>boolean</b>, <b>null</b>, <b>number</b>, <b>object</b>, <b>string</b> o <b>undefined</b></td><td>Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto <b>AsyncResult</b> sin sufrir modificaciones.</td><td></td></tr></table>|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **getSelectedDataAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Tener acceso a los valores de la selección actual, que se devuelven en la estructura de datos o en el formato que haya especificado con el parámetro _coercionType_ (consulte **Comentarios** para obtener más información sobre la coerción de datos).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **objeto** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## <a name="remarks"></a>Comentarios

En su complemento de panel de tareas o de contenido, use el método **getSelectedDataAsync** para escribir un script que lea los datos de la selección del usuario en un documento, hoja de cálculo, presentación o proyecto. Por ejemplo, cuando un usuario selecciona contenido en un documento de Word, puede usar el método **getSelectedDataAsync** para leer esa selección y luego enviarla a un servicio web como una consulta o alguna otra operación.

Una vez leída la selección, también puede usar los métodos [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) y [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) del objeto **Document** para [volver a escribir en la selección o agregar un controlador de eventos](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md) para detectar si el usuario cambia la selección.

El método **getSelectedDataAsync** puede leer de la selección siempre y cuando esté activo. En los complementos de Word y Excel, si necesita realizar una asociación persistente para leer y escribir en la selección del usuario, use en cambio el método [Bindings.addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) para [enlazar con esa selección](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).

Use el parámetro _coercionType_ del método **getSelectedDataAsync** para especificar el formato o la estructura de datos de los datos seleccionados que se están leyendo.



|**coercionType_ especificado_**|**Datos devueltos**|**Soporte para la aplicación host de Office**|
|:-----|:-----|:-----|
|**Office.CoercionType.Text** o `"text"`|Una cadena.|Una cadena.<br/><br/> **Nota**: en Excel, aunque se seleccione un subconjunto de una celda, se devuelve todo el contenido de la celda.|
|**Office.CoercionType.Matrix** o `"matrix"`|Una matriz de matrices. Por ejemplo, ` [['a','b'], ['c','d']]` para una selección de dos filas en dos columnas.|Word y Excel.|
|**Office.CoercionType.Table** o `"table"`|Un objeto [TableData](../../reference/shared/tabledata.md) para leer una tabla con encabezados.|Word y Excel.|
|**Office.CoercionType.Html** o `"html"`|En formato HTML.|En formato HTML.|
|**Office.CoercionType.Ooxml** o `"ooxml"`|En formato Open Office XML (OpenXML).|En formato HTML.<br/><br/> **Consejo**: Al desarrollar el código del complemento, puede usar el `"ooxml"`_coercionType_ del método **getSelectedDataAsync** para ver cómo el contenido seleccionado en un documento de Word se define como etiquetas de OpenXML. A continuación, use estas etiquetas en el parámetro de datos del método [Document.setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) para escribir contenido con este formato o esta estructura en un documento. Por ejemplo, puede [insertar una imagen en un documento](http://blogs.msdn.com/b/officeapps/archive/2012/10/26/inserting-images-with-apps-for-office.aspx) como OpenXML.|
|**Office.CoercionType.SlideRange** o "slideRange"|Un objeto JSON que contiene una matriz denominada "slides" (diapositivas), que a su vez contiene los identificadores, los títulos y los índices de las diapositivas seleccionadas.  **Nota:** Para seleccionar más de una diapositiva, el usuario debe editar la presentación en la vista **Normal**, **Vista esquema** o **Clasificador de diapositivas**. Además, este método no es compatible con **Vista Patrón**. Por ejemplo, `{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` para una selección de dos diapositivas.|Solo en PowerPoint.|
Si la estructura de datos de la selección no coincide con el _coercionType_ especificado, el método **getSelectedDataAsync** intentará convertir los datos en ese tipo o estructura. Si la selección no se puede convertir en el **Office.CoercionType** especificado, la propiedad **AsyncResult.status** devolverá `"failed"`.


## <a name="example"></a>Ejemplo

Para leer el valor de la selección actual, necesita escribir una función de respuesta que lea la selección. En el siguiente ejemplo se muestra cómo hacerlo:


-  **Pase una función de devolución de llamada anónima** que lea el valor de la selección actual al parámetro _callback_ del método **getSelectedDataAsync**.
    
-  **Lea la selección** como texto, sin formato y no filtrado.
    
-  **Mostrar el valor** en la página del complemento.
    

```js
function getText() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
        { valueFormat: "unformatted", filterType: "all" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            } 
            else {
                // Get selected data.
                var dataValue = asyncResult.value; 
                write('Selected data is ' + dataValue);
            }            
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
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
|**PowerPoint**|v|v|v|
|**Project**|v|||
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|Selección|
|**Nivel de permisos mínimo**|[ReadDocument (ReadAllDocument obligatorio para obtener Office Open XML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint Online.|
|1.1| En Word Online se ha agregado compatibilidad para **Office.CoercionType.Matrix** y **Office.CoercionType.Table** como parámetro _coercionType_.|
|1.1|En Excel, PowerPoint y Word en Office para iPad, se ha agregado el mismo nivel de compatibilidad que para Excel, PowerPoint y Word en el escritorio de Windows.|
|1.1| En Word Online se ha agregado compatibilidad para **Office.CoercionType.Text** como el parámetro _coercionType_.|
|1.1|En los complementos de contenido para PowerPoint puede obtener los identificadores, los títulos y los índices del rango de diapositivas seleccionado pasando **Office.CoercionType.SlideRange** como parámetro _coercionType_ del método **getSelectedDataAsync**. Consulte el tema relacionado con el método [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) para ver un ejemplo de cómo usar este valor para ir a la diapositiva que tiene seleccionada actualmente.|
|1.0|Agregado|
