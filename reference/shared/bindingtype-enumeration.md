
# Enumeración BindingType
 Especifica el tipo de objeto de enlace que se debería devolver.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Modificado por última vez**|1.1|

```
Office.BindingType
```


## Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.BindingType.Matrix|"matrix"|Datos tabulares sin fila de encabezado. Se devuelven los datos como una matriz de matrices; por ejemplo, con esta forma: ` [[row1column1, row1column2],[row2column1, row2column2]]`|
|Office.BindingType.Table|"table"|Datos tabuladores con una fila de encabezado. Se devuelven los datos como un objeto [TableData](../../reference/shared/tabledata.md).|
|Office.BindingType.Text|"text"|Texto sin formato. Se devuelven los datos con una sucesión de caracteres.|

## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta enumeración es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|v|||
|**Excel**|v|v|v|
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
|1.1|Se ha agregado compatibilidad para enlazar datos de tabla en los complementos para Access.|
|1.0|Agregado.|
