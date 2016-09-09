
# Enumeración InitializationReason
Especifica si el complemento se acaba de insertar o si se encontraba en el documento con anterioridad. 

|||
|:-----|:-----|
|**Hosts:**|Excel, Project y Word|
|**Agregado en**|1,0|

```
Office.InitializationReason
```


## Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.InitializationReason.Inserted|"inserted"|El complemento se acaba de insertar en el documento.|
|Office.InitializationReason.DocumentOpened|"documentOpened"|El complemento ya forma parte del documento abierto.|

## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta enumeración es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Project**|v|||
|**Word**|v||v|

|||
|:-----|:-----|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.0|Agregado|
