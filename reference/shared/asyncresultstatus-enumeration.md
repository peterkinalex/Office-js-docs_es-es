
# Enumeración AsyncResultStatus
Especifica el resultado de una llamada asíncrona. 

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project y Word|
|**Modificado por última vez en**|1.1|

```
Office.AsyncResultStatus
```


## Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.AsyncResultStatus.Succeeded|"succeeded"|La llamada ha sido correcta.|
|Office.AsyncResultStatus.Failed|"failed"|La llamada ha fallado.|

## Comentarios

Devuelto por la propiedad [status](../../reference/shared/asyncresult.status.md) del objeto [AsyncResult](../../reference/shared/asyncresult.md).


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta enumeración es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.


Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|**OWA para dispositivos**|**Office para Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|v|||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Project**|v|||||
|**Word**|v||v|||

|||
|:-----|:-----|
|**Tipos de complementos**|Contenido, panel de tareas y Outlook|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para los complementos para Access.|
|1.0|Agregado|
