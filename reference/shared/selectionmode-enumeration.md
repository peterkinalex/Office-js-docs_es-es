
# Enumeración SelectionMode
Especifica si se va a seleccionar (resaltar) la ubicación a la que se va a dirigir (al usar el método [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md)).

|||
|:-----|:-----|
|**Incorporado en la versión de Office.js**|1.1|

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint y Word|
|**Agregado en**|1.1|



```
Office.SelectionMode
```


## Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.SelectionMode.Selected|"selected"|Se seleccionará (resaltará) la ubicación.|
|Office.SelectionMode.None|"none"|El cursor se mueve al inicio de la ubicación.|

## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**PowerPoint**|v|||
|**Word**|v||v|

|||
|:-----|:-----|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Agregado|
