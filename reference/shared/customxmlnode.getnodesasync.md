
# <a name="customxmlnode.getnodesasync-method"></a>Método CustomXmlNode.getNodesAsync
Obtiene los nodos como una matriz de objetos **CustomXMLNode** que coinciden con la expresión XPath relativa.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Agregado en**|1.1|

```js
customXmlNodeObj.getNodesAsync(xPath [, options], callback);
```


## <a name="parameters"></a>Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _xPath_|**string**|Una expresión XPath que especifica los nodos que se deben obtener. Obligatorio.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **getNodesAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Obtener acceso a una **matriz** de objetos [CustomXMLNode](../../reference/shared/customxmlnode.customxmlnode.md) que representan los nodos especificados por la expresión XPath transferida al parámetro _xPath_.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **objeto** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## <a name="example"></a>Ejemplo




```js
function showXmlChildNodes() {
    Office.context.document.customXmlParts.getByIdAsync("{3BC85265-09D6-4205-B665-8EB239A8B9A1}", function (result) {
        var xmlPart = result.value;
        xmlPart.getNodesAsync('*', function (nodeResults) {
            for (i = 0; i < nodeResults.value.length; i++) {
                var node = nodeResults.value[i];
                node.getNodesAsync('*', function (nodeResults) {
                    write(nodeResults.value.length + " childNodes");
                });
            }
        });
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


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|CustomXmlParts|
|**Nivel de permisos mínimo**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word en Office para iPad.|
|1.0|Agregado|
