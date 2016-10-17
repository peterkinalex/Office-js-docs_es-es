
# <a name="document.customxmlparts-property"></a>Propiedad Document.customXmlParts
Obtiene un objeto que representa los elementos XML personalizados del documento.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Agregado en**|1.1|

```js
var xmlParts = Office.context.document.customXmlParts;
```


## <a name="return-value"></a>Valor devuelto

Un objeto [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md).


## <a name="example"></a>Ejemplo




```js
function getCustomXmlParts(){
    Office.context.document.customXmlParts.getByNamespaceAsync('http://tempuri.org', function (asyncResult) {
        write('Retrieved ' + asyncResult.value.length + ' custom XML parts');
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|v||v|

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word en Office para iPad.|
|1.0|Agregado|
