
# Método CustomXmlNode.setTextAsync
Obtiene el texto de un nodo XML de forma asíncrona en un elemento XML personalizado.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Agregado en**|1.2|

```
customXmlNodeObj.setTextAsync(text, [asyncContext,]callback(asyncResult);
```


## Parámetros



|**Nombre**|**Tipo**|**Descripción**|
|:-----|:-----|:-----|
| _text_|**string**|Necesario. El valor de texto del nodo XML.|
| _asyncContext_|**object**|Opcional. Un objeto definido por el usuario que está disponible en la propiedad asyncContext del objeto [AsyncResult](../../reference/shared/asyncresult.md). Use esta opción para proporcionar un objeto o valor para el **AsyncResult** si la devolución de llamada es una función con nombre.|
| _callback_|**object**|Opcional. Una función que se invoca al devolver una llamada, cuyo único parámetro es de tipo **AsyncResult**.|

## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **setTextAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|No se usa.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Indica si la operación se ha efectuado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **object** definidos por el usuario si ha remitido uno como parámetro _asyncContext_. Esta propiedad devuelve «undefined» si no se ha establecido _asyncContext_.|

## Ejemplo

Obtenga información sobre cómo establecer el valor de texto de un nodo en un elemento XML personalizado.


```js
// Get the built-in core properties XML part by using its ID. This results in a call to Word.
Office.context.document.customXmlParts.getByIdAsync("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}", function (getByIdAsyncResult) {
    
    // Access the XML part.
    var xmlPart = getByIdAsyncResult.value;
    
    // Add namespaces to the namespace manager. These two calls result in two calls to Word.
    xmlPart.namespaceManager.addNamespaceAsync('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', function () {
        xmlPart.namespaceManager.addNamespaceAsync('dc', 'http://purl.org/dc/elements/1.1/', function () {

            // Get XML nodes by using an Xpath expression. This results in a call to the host.
            xmlPart.getNodesAsync("/cp:coreProperties/dc:subject", function (getNodesAsyncResult) {
                
                // Get the first node returned by using the Xpath expression. This will be the subject element in this example.
                var subjectNode = getNodesAsyncResult.value[0];
                
                // Set the text value of the subject node and use the asyncContext. This results in a call to the host. 
                // The results are logged to the browser console. 
                subjectNode.setTextAsync("newSubject", {asyncContext: "StateNormal"}, function (setTextAsyncResult) {
                   console.log("The status of the call: " + setTextAsyncResult.status);
                   console.log("The asyncContext value = " + setTextAsyncResult.asyncContext);
                });
            });
        });
    });
});
```


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|CustomXmlParts|
|**Nivel de permisos mínimo**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado setTextAsync.|
