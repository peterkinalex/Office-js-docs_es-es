
# <a name="matrixbinding-object"></a>Objeto MatrixBinding
Representa un enlace en dos dimensiones de filas y columnas. 

|||
|:-----|:-----|
|**Hosts:**|Excel y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings|
|**Modificado por última vez en Selección**|1.1|

```
MatrixBinding
```


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[columnCount](../../reference/shared/binding.matrixbinding.columncount.md)|Obtiene la cantidad de columnas de la estructura de datos matriz en forma de valor entero.|
|[rowCount](../../reference/shared/binding.matrixbinding.rowcount.md)|Obtiene el número de filas de la estructura de datos de matriz como un valor entero.|

## <a name="remarks"></a>Comentarios

El objeto **MatrixBinding** hereda las propiedades [id](../../reference/shared/binding.id.md) y [type](../../reference/shared/binding.type.md), y los métodos [getDataAsync](../../reference/shared/binding.getdataasync.md) y [setDataAsync](../../reference/shared/binding.setdataasync.md) del objeto [Binding](../../reference/shared/binding.md).


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|MatrixBindings|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.0|Agregado|
