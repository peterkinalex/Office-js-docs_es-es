
# Enumeración DocumentMode
Especifica si el documento de la aplicación asociada es de solo lectura o de lectura y escritura. 

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint, Project y Word|
|**Agregado en**|1.1|

```
Office.DocumentMode
```


## Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.DocumentMode.ReadOnly|"readOnly"|El documento es de solo lectura.|
|Office.DocumentMode.ReadWrite|"readWrite"|El documento es de lectura y escritura.|

## Comentarios

Devuelto por la propiedad **mode** del objeto [Document](../../reference/shared/document.md).


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta enumeración es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Project**|v|||
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
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.0|Agregado|
