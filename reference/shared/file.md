
# <a name="file-object"></a>Objeto File
Representa el archivo de documento asociado a un complemento de Office.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Archivo|
|**Modificado por última vez en**|1.1|

```
file
```


## <a name="members"></a>Miembros


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|**[size](../../reference/shared/file.size.md)**|Obtiene el tamaño del archivo de documento en bytes.|
|**[sliceCount](../../reference/shared/file.slicecount.md)**|Obtiene el número de segmentos en los que está dividido el archivo.|

**Métodos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|**[closeAsync](../../reference/shared/file.closeasync.md)**|Cierra el archivo de documento.|
|**[getSliceAsync](../../reference/shared/file.getsliceasync.md)**|Devuelve el segmento especificado.|

## <a name="remarks"></a>Comentarios

Obtiene acceso al objeto **File** con la propiedad [AsyncResult.value](../../reference/shared/asyncresult.value.md) de la función de devolución de llamada que se remite al método [Document.getFileAsync](../../reference/shared/document.getfileasync.md).


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||||
|:-----|:-----|:-----|:-----|
||Office para escritorio de Windows|Office Online (en el explorador)|Office para iPad|
|**PowerPoint**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible en el conjunto de requisitos**|Archivo|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint y Word en Office para iPad.|
|1.0|Agregado|
