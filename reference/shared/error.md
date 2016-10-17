
# <a name="error-object"></a>Objeto Error
Proporciona información específica sobre un error que se produjo durante una operación de datos asincrónica.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project y Word|
|**Modificado por última vez en**|1.1|

```
asyncResult.error
```


## <a name="members"></a>Miembros


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[code](../../reference/shared/error.code.md)|Obtiene el código numérico del error.|
|[name](../../reference/shared/error.name.md)|Obtiene el nombre del error.|
|[message](../../reference/shared/error.message.md)|Obtiene una descripción detallada del error.|

## <a name="remarks"></a>Comentarios

Puede obtener acceso al objeto **Error** desde el objeto [AsyncResult](../../reference/shared/asyncresult.md). Este último objeto se devuelve en la función que se ha remitido como argumento _callback_ de una operación de datos asincrónicos (por ejemplo, el método [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) del objeto **Document**).


## <a name="example"></a>Ejemplo

En el ejemplo siguiente se usa el método **setSelectedDataAsync** para establecer el texto seleccionado en "Hello World!" y, si se produce un error, mostrar los valores de las propiedades **name** y **message** del objeto **Error**.


```js
function setText() {

    Office.context.document.setSelectedDataAsync("Hello World!", {},
        function (asyncResult) {
            if (asyncResult.status === "failed")
            var err = asyncResult.error; 
                write(err.name + ": " + err.message);
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

||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|**OWA para dispositivos**|**Outlook para Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||v||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Project**|v|||||
|**Word**|v|v|v|||

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas y Outlook|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad con complementos de contenido para Access.|
|1.0|Agregado|
