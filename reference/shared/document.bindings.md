
# <a name="document.bindings-property"></a>Propiedad Document.bindings
Obtiene un objeto que proporciona acceso a los enlaces que se han definido en el documento.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Modificado por última vez en**|1.1|

```js
var docBindings = Office.context.document.bindings;
```


## <a name="return-value"></a>Valor devuelto

Un objeto [Bindings](../../reference/shared/bindings.bindings.md).


## <a name="example"></a>Ejemplo




```js
function displayAllBindings() {
    Office.context.document.bindings.getAllAsync(function (asyncResult) {
        var bindingString = '';
        for (var i in asyncResult.value) {
            bindingString += asyncResult.value[i].id + '\n';
        }
        write('Existing bindings: ' + bindingString);
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
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad





****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para la escritura de datos de tabla en complementos para Access.|
|1.0|Agregado|
