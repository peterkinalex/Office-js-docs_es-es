
# Propiedad Bindings.document
Obtiene un objeto **Document** que representa el documento asociado a este conjunto de enlaces.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Modificado por última vez en**|1.1|

```
var docObj = bindingsObj.document;
```


## Valor devuelto

Un objeto [Document](../../reference/shared/bindings.document.md).


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

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
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Se ha agregado acceso a un objeto **Document** que representa la base de datos actual de Access en los complementos de contenido para Access.|
|1.0|Agregado|
