
# Método CustomXmlNode.getTextAsync
Obtiene el texto de un nodo XML de forma asíncrona en un elemento XML personalizado.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Agregado en**|1.2|

```js
customXmlNodeObj.getTextAsync([asyncContext,]callback(asyncResult);
```


## Parámetros



|**Nombre**|**Tipo**|**Descripción**|
|:-----|:-----|:-----|
| _asyncContext_|**object**|Opcional. Un objeto definido por el usuario que está disponible en la propiedad asyncContext del objeto [AsyncResult](../../reference/shared/asyncresult.md). Use esta opción para proporcionar un objeto o valor para el **AsyncResult** si la devolución de llamada es una función con nombre.|
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.|

## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **getTextAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Tener acceso a una **string** que contiene el texto interno de los nodos referenciados.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Indica si la operación se ha efectuado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **object** definidos por el usuario si ha remitido uno como parámetro _asyncContext_. Esta propiedad devuelve «undefined» si no se ha establecido _asyncContext_.|

## Ejemplo

Obtenga información sobre cómo obtener el valor de texto de un nodo en un elemento XML personalizado.


```js
// Get the built-in core properties XML part by using its ID. This results in a call to Word.
Office.context.document.customXmlParts.getByIdAsync("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}", function (getByIdAsyncResult) {
    
    // Access the XML part.
    var xmlPart = getByIdAsyncResult.value;
    
    // Add namespaces to the namespace manager. These two calls result in two calls to Word.
    xmlPart.namespaceManager.addNamespaceAsync('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', function () {
        xmlPart.namespaceManager.addNamespaceAsync('dc', 'http://purl.org/dc/elements/1.1/', function () {

            // Get XML nodes by using an Xpath expression. This results in a call to Word.
            xmlPart.getNodesAsync("/cp:coreProperties/dc:title", function (getNodesAsyncResult) {
                
                // Get the first node returned by using the Xpath expression. 
                var node = getNodesAsyncResult.value[0];
                
                // Get the text value of the node and use the asyncContext. This results in a call to Word. 
                // The results are logged to the browser console.
                node.getTextAsync({asyncContext: "StateNormal"}, function (getTextAsyncResult) {
                   console.log("Text of the title element = " + getTextAsyncResult.value;
                   console.log("The asyncContext value = " + getTextAsyncResult.asyncContext;
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
|1.1|Se ha agregado getTextAsync.|
