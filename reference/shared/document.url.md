
# Propiedad Document.url
Obtiene la dirección URL del documento que se encuentra abierto actualmente en la aplicación host.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Project y Word|
|**Modificado por última vez en**|1.1|

```
var docUrl = Office.context.document.url;
```


## Valor devuelto

La dirección URL del documento. Devuelve **null** si la dirección URL no se encuentra disponible.


## Comentarios

 **Importante:** la propiedad **url** devuelve información que puede contener información de identificación personal (PII) en el nombre del documento y la ubicación donde se almacena. En caso de que deba almacenar o transmitir esta información, hágalo en formato cifrado.


## Ejemplo




```
function displayDocumentUrl() {
    write(Office.context.document.url);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Project**|v|||
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad





****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word Online.|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad con complementos de contenido para Access.|
|1.0|Agregado|
