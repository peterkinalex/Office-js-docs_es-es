
# <a name="tabledata-object"></a>Objeto TableData
Representa los datos en una tabla o un [TableBinding](../../reference/shared/binding.tablebinding.md).

|||
|:-----|:-----|
|**Hosts:**|Excel y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Agregado en**|1.1|

```
TableData
```

## <a name="members"></a>Miembros


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[headers](../../reference/shared/tabledata.headers.md)|Obtiene o establece los encabezados de la tabla.|
|[rows](../../reference/shared/tabledata.rows.md)|Obtiene o establece las filas de la tabla.|

## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|TableBindings|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word Online.|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.0|Agregado|
