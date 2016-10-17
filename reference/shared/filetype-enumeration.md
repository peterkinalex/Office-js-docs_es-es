
# <a name="filetype-enumeration"></a>Enumeración FileType
Especifica el formato en el que deben devolverse los documentos.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint y Word|
|**Modificado por última vez en**|1.1|

```js
Office.FileType
```


## <a name="members"></a>Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.FileType.Compressed|"compressed"|Devuelve el documento completo (.pptx o .docx) en el formato Office Open XML (OOXML) como una matriz de bytes.|
|Office.FileType.Pdf|"pdf"|Devuelve todo el documento en formato PDF como matriz de bytes.|
|Office.FileType.Text|"text"|Devuelve solo el texto del documento como una **string** (solo en Word).|

## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta enumeración es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para guardar como PDF.|
|1.0|Agregado|
