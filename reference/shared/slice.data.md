
# <a name="slice.data-property"></a>Propiedad Slice.data
Obtiene los datos sin procesar del segmento del archivo.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Archivo|
|**Modificado por última vez en**|1.1|

```
var sliceData = slice.data;
```


## <a name="return-value"></a>Valor devuelto

Los datos sin procesar del segmento del archivo en el formato **Office.FileType.Text** ("texto") u **Office.FileType.Compressed** ("comprimido") como especifica el parámetro _fileType_ de la llamada al método [Document.getFileAsync](../../reference/shared/document.getfileasync.md).


## <a name="remarks"></a>Comentarios

Los archivos en formato "comprimido" devolverán una matriz de bytes que puede transformarse en una cadena con codificación Base 64 si es necesario.


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|


|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|Archivo|
|**Nivel de permisos mínimo**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint y Word en Office para iPad.|
|1.0|Agregado|
