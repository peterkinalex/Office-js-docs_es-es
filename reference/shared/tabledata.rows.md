
# Propiedad TableData.rows
Obtiene o establece las filas de la tabla.

|||
|:-----|:-----|
|**Hosts:**|Excel y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Agregado en**|1.1|

```
var myRows = tableBindingObj.rows;
```


## Valor devuelto

Devuelve una matriz de matrices que contiene los datos de la tabla. Si no hay ninguna fila, devuelve una **array**`[]` vacía.


## Observaciones

Para especificar las filas, debe indicar una matriz de matrices que corresponda a la estructura de la tabla. Por ejemplo, para especificar dos filas de valores **string** en una tabla de dos columnas, establezca la propiedad **rows** en ` [['a', 'b'], ['c', 'd']]`.

Si especifica **null** para la propiedad **rows** (o la deja vacía al construir un objeto **TableData**), al ejecutarse el código se producirán los siguientes resultados:


- Si inserta una nueva tabla, se insertará una fila en blanco.
    
- Si sobrescribe o actualiza una tabla existente, las filas existentes no se modificarán.
    

## Ejemplo

En el ejemplo siguiente se crea una tabla de una sola columna con un encabezado y tres filas.


```js
function createTableData() {
    var tableData = new Office.TableData();
    tableData.headers = [['header1']];
    tableData.rows = [['row1'], ['row2'], ['row3']];
    return tableData;
}
```


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v|v|v|


|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|TableBindings|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word Online.|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.0|Agregado|
