
# <a name="tablebinding-object"></a>Objeto TableBinding
Representa un enlace en dos dimensiones de filas y columnas, que puede llevar o no encabezados.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Modificado por última vez en Selección**|1.1|

```
TableBinding
```


## <a name="members"></a>Miembros


**Propiedades**


|**Nombre**|**Descripción**|**Actualizaciones de Office.js v1.1**|
|:-----|:-----|:-----|
|[columnCount](../../reference/shared/binding.tablebinding.columncount.md)|Obtiene el número de columnas que hay en el objeto **TableBinding** especificado.|Se agregó compatibilidad para el enlace de tablas en los complementos de contenido para Access.|
|[hasHeaders](../../reference/shared/binding.tablebinding.hasheaders.md)|Si el objeto **TableBinding** especificado tiene encabezados, devolverá verdadero. De lo contrario, devolverá falso.|Se agregó compatibilidad para el enlace de tablas en los complementos de contenido para Access.|
|[rowCount](../../reference/shared/binding.tablebinding.rowcount.md)|El número de filas existentes en el objeto **TableBinding** especificado.|En los complementos de contenido para Access, siempre devuelve -1 por motivos de rendimiento.|

**Métodos**


|**Nombre**|**Descripción**|**Actualizaciones de Office.js v1.1**|
|:-----|:-----|:-----|
|[addColumnsAsync](../../reference/shared/binding.tablebinding.addcolumnsasync.md)|Agrega columnas y valores a una tabla.||
|[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|Agrega filas y valores a una tabla.|Se agregó compatibilidad para el enlace de tablas en los complementos de contenido para Access.|
|[clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md)|Borra el formato en la tabla enlazada.|Nuevo en Office.js v1.1 para los complementos para Excel.|
|[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|Elimina de la tabla todas las filas que no sean encabezados y sus valores, y cambia de forma adecuada a la aplicación host.|Se agregó compatibilidad para el enlace de tablas en los complementos de contenido para Access.|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|Escribe datos en la sección enlazada del documento que representa el objeto de enlace que se ha especificado.|<ul><li>Se ha agregado compatibilidad para el enlace de tablas en los complementos de contenido para Access.</li><li>Se agregó compatibilidad para establecer el formato al escribir datos en tablas enlazadas en los complementos para Excel.</li></ul>|
|[setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|Establece el formato de tabla y celda en los datos y los elementos especificados de la tabla enlazada.|Puede establecer el formato de tabla en los complementos para Excel.|
|[setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md)|Actualiza las opciones de formato de tabla en la tabla enlazada.|Puede establecer el formato de tabla en los complementos para Excel.|

## <a name="remarks"></a>Comentarios

El objeto **TableBinding** hereda las propiedades [id](../../reference/shared/binding.id.md) y [type](../../reference/shared/binding.type.md) y los métodos [getDataAsync](../../reference/shared/binding.getdataasync.md) y [setDataAsync](../../reference/shared/binding.setdataasync.md) del objeto abstracto [Binding](../../reference/shared/binding.md).

Cuando establezca un enlace de tabla en Excel, se incluirán automáticamente en el enlace todas las filas nuevas que agregue un usuario a dicha tabla (y aumentará **rowCount**).


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|TableBindings|
|**Nivel de permisos mínimo**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para [establecer el formato al insertar tablas](../../docs/excel/format-tables-in-add-ins-for-excel.md) en Excel.|
|1.1|Se ha agregado compatibilidad para los complementos para Access.|
|1.0|Agregado|
