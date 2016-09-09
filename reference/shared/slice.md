
# Objeto Slice
Representa un segmento de un archivo de documento.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Archivo|
|**Modificado por última vez en**|1.1|

```
slice
```


## Miembros


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|**[data](../../reference/shared/slice.data.md)**|Obtiene los datos sin procesar del segmento del archivo.|
|**[index](../../reference/shared/slice.index.md)**|Obtiene el índice del segmento de archivos.|
|**[size](../../reference/shared/slice.size.md)**|Obtiene el tamaño del segmento en bytes.|

## Comentarios

Se obtiene acceso al objeto **Slice** con el método [File.getSliceAsync](../../reference/shared/file.getsliceasync.md).


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|


|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|Archivo|
|**Nivel de permisos mínimo**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint y Word en Office para iPad.|
|1.0|Agregado|
