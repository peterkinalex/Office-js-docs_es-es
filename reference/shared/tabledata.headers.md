
# <a name="tabledata.headers-property"></a>Propiedad TableData.headers
Obtiene o establece los encabezados de la tabla.

|||
|:-----|:-----|
|**Hosts:**|Excel y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Modificado por última vez en**|1.1|

```
var hasHeaders = tableBindingObj.headers;
```


## <a name="return-value"></a>Valor devuelto

 **true** si la tabla tiene encabezados; si no, **false**. 


## <a name="remarks"></a>Observaciones

Para especificar encabezados, debe especificar una matriz de matrices que se corresponda con la estructura de la tabla. Por ejemplo, para especificar los encabezados de una tabla de dos columnas, debe establecer la propiedad **header** en ` [['header1', 'header2']]`.

Si especifica **null** para la propiedad **headers** (o la deja vacía al construir un objeto **TableData**), al ejecutarse el código se producirán los siguientes resultados:


- Si inserta una tabla nueva, se crearán los encabezados de columna predeterminados para la tabla.
    
- Si sobrescribe o actualiza una tabla existente, los encabezados existentes no se modificarán.
    

## <a name="example"></a>Ejemplo

En el ejemplo siguiente se crea una tabla de una sola columna con un encabezado y tres filas.


```js
function createTableData() {
    var tableData = new Office.TableData();
    tableData.headers = [['header1']];
    tableData.rows = [['row1'], ['row2'], ['row3']];
    return tableData;
}

```


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|TableBindings|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word Online.|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.0|Agregado|
