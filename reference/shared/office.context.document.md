
# Propiedad Context.document
Obtiene un objeto que representa el documento con el que el complemento está interactuando.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project y Word|
|**Modificado por última vez en**|1.1|

```js
var _document = Office.context.document;
```


## Valor devuelto

Un objeto [Document](../../reference/shared/document.md).


## Comentarios

El complemento puede usar la propiedad **document** para obtener acceso a la API e interactuar con el contenido de documentos, libros, presentaciones, proyectos y bases de datos (en las aplicaciones web de Access).


## Ejemplo




```js
// Extension initialization code.
var _document;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, code specific to the add-in can run.
    // Initialize instance variables to access API objects.
    _document = Office.context.document;
    });
}

```


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Project**|v|||
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para **Office.context.document** con el fin de obtener acceso a la base de datos de los complementos de contenido para Access.|
|1.0|Agregado|
