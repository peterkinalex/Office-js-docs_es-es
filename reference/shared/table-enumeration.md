
# <a name="table-enumeration"></a>Enumeración Table
Especifica los valores enumerados de la propiedad `cells:` en el parámetro _cellFormat_ de los [métodos de formato de tabla](../../docs/excel/format-tables-in-add-ins-for-excel.md).

|||
|:-----|:-----|
|**Hosts:**|Excel|
|**Agregado**|1.1|

```
Office.Table
```

## <a name="members"></a>Miembros


**Valores**


|**Enumeración**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|Office.Table.All|"all"|Toda la tabla, incluidos encabezados de columnas, datos y totales (si existen).|
|Office.Table.Data|"data"|Solo los datos (sin encabezados ni totales).|
|Office.Table.Headers|"headers"|Solo la fila de encabezado.|

## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica las enumeraciones compatibles con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta enumeración.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|

|||
|:-----|:-----|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel en Office para iPad.|
|1.1|Agregado|
