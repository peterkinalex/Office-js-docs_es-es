
# Propiedad Context.touchEnabled
Obtiene información sobre si el complemento se está ejecutando en una aplicación host de Office con funcionalidad táctil.

|||
|:-----|:-----|
|**Hosts:**|Excel y Word|
|**Modificado por última vez en**|1.1|

```
var isTouchEnabled = Office.context.touchEnabled;
```


## Valor devuelto

Devuelve **True** si el complemento se está ejecutando en un dispositivo táctil, como un iPad; de lo contrario, devuelve **False**.


## Comentarios

Use la propiedad **touchEnabled** para determinar si el complemento se está ejecutando en un dispositivo táctil y, si fuera necesario, ajuste el tipo de controles y el tamaño y el espaciado de los elementos de la IU del complemento para permitir las interacciones táctiles.


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office para iPad**|
|:-----|:-----|
|**Excel**|v|
|**PowerPoint**|v|
|**Word**|v|

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
|1.1|Agregado.|
