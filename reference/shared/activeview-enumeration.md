
# Enumeración ActiveView
Especifica el estado de la vista activa del documento (por ejemplo, si el usuario puede editar o no el documento).

|||
|:-----|:-----|
|**Incorporado en la versión de Office.js**|1.1|

|||
|:-----|:-----|
|**Hosts:**|PowerPoint|
|**Agregado en**|1.1|



```
Office.ActiveView
```


## Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.ActiveView.Read|"read"|La vista activa de la aplicación host solo permite al usuario leer el contenido del documento.|
|Office.ActiveView.Edit|"edit"|La vista activa de la aplicación host permite al usuario editar el contenido del documento.|

## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta enumeración es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|v|v|v|

|||
|:-----|:-----|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint en Office para iPad.|
|1.1|Agregado|
