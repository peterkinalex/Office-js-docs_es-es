
# Propiedad Error.message
Obtiene una descripción detallada del error.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project y Word|
|**Modificado por última vez en Selección**|1.1|

```
var errMessage = asyncResult.error.message;
```


## Valor devuelto

La descripción del error como una **string**.


## Comentarios

Al objeto **Error** y a sus propiedades se obtiene acceso desde el objeto [AsyncResult](../../reference/shared/asyncresult.md) que se devuelve en la función transferida como el argumento _callback_ de una acción de datos asíncrona.


## Ejemplo

Para que se devuelva un error, seleccione una tabla o una matriz y después llame a la función `setText`.


```js
function setText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            if (asyncResult.status === "failed")
                var error = asyncResult.error;
            write(error.name + ": " + error.message);
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Detalles de compatibilidad


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

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint Online.|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad con complementos de contenido para Access.|
|1.0|Agregado|
