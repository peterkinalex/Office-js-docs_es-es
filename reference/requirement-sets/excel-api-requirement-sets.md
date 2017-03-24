# <a name="excel-javascript-api-requirement-sets"></a>Conjuntos de requisitos de la API de JavaScript de Excel

Los conjuntos de requisitos son grupos de miembros de la API con nombre. Los complementos de Office usan los conjuntos de requisitos especificados en el manifiesto o usan una comprobación en tiempo de ejecución para determinar si un host de Office admite las API necesarias para el complemento. Para obtener más información, consulte [Specify Office hosts and API requirements (Especificar hosts de Office y requisitos de la API)](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Los complementos de Excel se ejecutan en varias versiones de Office, incluida Office 2016 para Windows, Office para iPad, Office para Mac y Office Online. En la siguiente tabla se enumeran los conjuntos de requisitos de Excel, las aplicaciones de host de Office que admiten ese conjunto de requisitos y el número o las versiones de compilación de esas aplicaciones.

> Para los conjuntos de requisitos que están marcados como *versión Beta*, utilice la versión especificada (o posterior) del software de Office y la biblioteca de la versión Beta de CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

> Las entradas que no se enumeran como *versión Beta* están generalmente disponibles y puede seguir utilizando la biblioteca de CDN de producción: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.

|  Conjunto de requisitos  |  Office 2016 para Windows*  |  Office 2016 para iPad  |  Office 2016 para Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| ExcelApi 1.5 **Beta**  | Versión 1702 (compilación por determinar) o posterior| Próximamente |  Próximamente| próximamente | Próximamente|
| ExcelApi 1.4 **Beta** | Versión 1702 (compilación por determinar) o posterior| Próximamente |  Próximamente| próximamente | Próximamente|
| ExcelApi 1.3  | Versión 1608 (compilación 7369.2055) o posterior| 1.27 o posterior |  15.27 o posterior| Septiembre de 2016 | Versión 1608 (compilación 7601.6800) o posterior|
| ExcelApi 1.2  | Versión 1601 (compilación 6741.2088) o posterior | 1.21 o posterior | 15.22 o posterior| Enero de 2016 ||
| ExcelApi 1.1  | Versión 1509 (compilación 4266.1001) o posterior | 1.19 o posterior | 15.20 o posterior| Enero de 2016 ||

> **Nota**: El número de compilación para Office 2016 instalado mediante MSI es el 16.0.4266.1001. Esta versión solo contiene el conjunto de requisitos de ExcelApi 1.1.

Para obtener más información sobre las versiones, números de compilación y Office Online Server, consulte:

- [Números de versión y compilación de las versiones del canal de actualización para los clientes de Office 365](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [¿Qué versión de Office estoy usando?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Dónde puede encontrar el número de versión y de compilación de una aplicación de cliente de Office 365](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- [Información general de Office Online Server](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API de Office
Para obtener información sobre los conjuntos de requisitos comunes de la API, consulte [Office common API requirement sets (Conjuntos de requisitos comunes de la API de Office)](office-add-in-requirement-sets.md).

## <a name="whats-new-in-excel-javascript-api-14"></a>Novedades de la API de JavaScript de Excel 1.4
Las siguientes son las nuevas incorporaciones a las API de JavaScript de Excel en el conjunto de requisitos 1.3.

### <a name="named-item-add-and-new-properties"></a>Agregar elementos con nombre y nuevas propiedades

Nuevas propiedades
* `comment`
* `scope`: elementos con ámbito de hoja de cálculo o libro.
* `worksheet`: devuelve la hoja de cálculo que tiene como ámbito el elemento con nombre.

Nuevos métodos
* `add(name: string, reference: Range or string, comment: string)`: agrega un nuevo nombre a la colección del ámbito especificado.
* `addFormulaLocal(name: string, formula: string, comment: string)`: agrega un nuevo nombre a la colección del ámbito especificado con la configuración regional del usuario para la fórmula.

### <a name="settings-api-in-in-excel-namespace"></a>API de configuración en el espacio de nombres de Excel

El objeto [Setting](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_1.4_OpenSpec/reference/excel/setting.md) representa un par clave-valor de una configuración que se conserva en el documento. Ahora hemos agregado unas API relacionadas con la configuración en el espacio de nombres de Excel. De esta forma, aunque la funcionalidad no es estrictamente nueva, resulta más fácil permanecer en la sintaxis de API por lotes basada en compromisos y se reduce la dependencia de la API común para tareas relacionadas con Excel.

Las API incluyen `getItem()` para obtener acceso a la configuración mediante la clave `add()` para agregar al libro el par clave-valor de configuración especificado.

### <a name="others"></a>Otros

* Establecer el nombre de la columna de tabla (la versión anterior solo permite la lectura).
* Agregar una columna al final de la tabla (la versión anterior permite cualquier lugar excepto el último).
* Agregar varias filas a una tabla de una sola vez (la versión anterior solo permite agregarlas de una en una).
* `range.getColumnsAfter(count: number)` y `range.getColumnsBefore(count: number)` para obtener un número determinado de columnas a la derecha o izquierda del objeto Range actual.
* Función para obtener un elemento o un objeto NULL: esta funcionalidad permite obtener el objeto mediante una clave. Si el objeto no existe, la propiedad isNullObject del objeto devuelto será "true". Esto permite a los desarrolladores comprobar si existe un objeto sin tener que utilizar el control de excepciones para controlarlo. Disponible para hojas de cálculo, elementos con nombre, enlaces, series de gráfico, etc.

`worksheet.GetItemOrNullObject()`

### <a name="suspend-calculation"></a>Suspender cálculo
Suspende el cálculo (application.suspendCalculationUntilNextSync()) hasta que se llama al siguiente "context.sync()". Una vez establecido, será responsabilidad del desarrollador actualizar el libro para asegurarse de que se propaguen las dependencias.

Además, hemos corregido el error al usar F9 para actualizar, que omitía las celdas desfasadas.

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:----|:----|:----|:----|
|[application](../excel/application.md)|_Método_ > [suspendCalculationUntilNextSync()](../excel/application.md#suspendcalculationuntilnextsync)|Suspende el cálculo hasta que se llama al siguiente "context.sync()". Una vez establecido, será responsabilidad del desarrollador actualizar el libro para asegurarse de que se propaguen las dependencias.|1.4|
|[bindingCollection](../excel/bindingcollection.md)|_Método_ > [getCount()](../excel/bindingcollection.md#getcount)|Obtiene el número de enlaces de la colección.|1.4|
|[bindingCollection](../excel/bindingcollection.md)|_Método_ > [getItemOrNullObject(id: cadena)](../excel/bindingcollection.md#getitemornullobjectid-string)|Obtiene un objeto de enlace por identificador. Si no existe el objeto de enlace, devolverá un objeto nulo.|1.4|
|[chartCollection](../excel/chartcollection.md)|_Método_ > [getCount()](../excel/chartcollection.md#getcount)|Devuelve el número de gráficos de la hoja de cálculo.|1.4|
|[chartCollection](../excel/chartcollection.md)|_Método_ > [getItemOrNullObject(name: string)](../excel/chartcollection.md#getitemornullobjectname-string)|Obtiene un gráfico mediante su nombre. Si hay varias tablas con el mismo nombre, se devolverá la primera.|1.4|
|[chartPointsCollection](../excel/chartpointscollection.md)|_Método_ > [getCount()](../excel/chartpointscollection.md#getcount)|Devuelve el número de puntos del gráfico de la serie.|1.4|
|[chartSeriesCollection](../excel/chartseriescollection.md)|_Método_ > [getCount()](../excel/chartseriescollection.md#getcount)|Devuelve el número de series incluidas en la colección.|1.4|
|[namedItem](../excel/nameditem.md)|_Propiedad_ > comment|Representa el comentario asociado a este nombre.|1.4|
|[namedItem](../excel/nameditem.md)|_Propiedad_ > scope|Indica si el nombre está en el ámbito del libro o de una hoja de cálculo específica. Solo lectura. Los valores posibles son: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](../excel/nameditem.md)|_Relación_ > worksheet|Devuelve la hoja de cálculo que tiene como ámbito el elemento con nombre. Se produce un error si el ámbito del elemento es el libro. Solo lectura.|1.4|
|[namedItem](../excel/nameditem.md)|_Relación_ > worksheetOrNullObject|Devuelve la hoja de cálculo que tiene como ámbito el elemento con nombre. Devuelve un objeto NULL si el ámbito del elemento es el libro. Solo lectura.|1.4|
|[namedItem](../excel/nameditem.md)|_Método_ > [delete()](../excel/nameditem.md#delete)|Elimina el nombre especificado.|1.4|
|[namedItem](../excel/nameditem.md)|_Método_ > [getRangeOrNullObject()](../excel/nameditem.md#getrangeornullobject)|Devuelve el objeto de rango asociado al nombre. Devuelve un objeto NULL si el tipo de elemento con nombre no es un rango.|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_Método_ > [add(name: string, reference: Range or string, comment: string)](../excel/nameditemcollection.md#addname-string-reference-range-or-string-comment-string)|Agrega un nuevo nombre a la colección del ámbito especificado.|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_Método_ > [addFormulaLocal(name: string, formula: string, comment: string)](../excel/nameditemcollection.md#addformulalocalname-string-formula-string-comment-string)|Agrega un nuevo nombre a la colección del ámbito especificado, empleando la configuración regional del usuario para la fórmula.|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_Método_ > [getCount()](../excel/nameditemcollection.md#getcount)|Obtiene el número de elementos con nombre de la colección.|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_Método_ > [getItemOrNullObject(name: string)](../excel/nameditemcollection.md#getitemornullobjectname-string)|Obtiene un objeto NamedItem mediante su nombre. Si no existe el objeto NamedItem, devolverá un objeto NULL.|1.4|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Método_ > [getCount()](../excel/pivottablecollection.md#getcount)|Obtiene el número de tablas dinámicas de una colección.|1.4|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Método_ > [getItemOrNullObject(name: string)](../excel/pivottablecollection.md#getitemornullobjectname-string)|Obtiene una tabla dinámica por nombre. Si no existe la tabla dinámica, devolverá un objeto NULL.|1.4|
|[range](../excel/range.md)|_Método_ > [getIntersectionOrNullObject(anotherRange: Range or string)](../excel/range.md#getintersectionornullobjectanotherrange-range-or-string)|Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados. Si no se encuentra ninguna intersección, se devolverá un objeto NULL.|1.4|
|[range](../excel/range.md)|_Método_ > [getUsedRangeOrNullObject(valuesOnly: bool)](../excel/range.md#getusedrangeornullobjectvaluesonly-bool)|Devuelve el rango usado del objeto de rango especificado. Si no hay ninguna celda usada dentro del rango, esta función devolverá un objeto NULL.|1.4|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_Método_ > [getCount()](../excel/rangeviewcollection.md#getcount)|Obtiene el número de objetos RangeView de la colección.|1.4|
|[setting](../excel/setting.md)|_Propiedad_ > key|Devuelve una clave que representa el identificador de la configuración. Solo lectura.|1.4|
|[setting](../excel/setting.md)|_Propiedad_ > value|Representa el valor almacenado para esta configuración.|1.4|
|[setting](../excel/setting.md)|_Método_ > [delete()](../excel/setting.md#delete)|Elimina la configuración.|1.4|
|[settingCollection](../excel/settingcollection.md)|_Propiedad_ > items|Una colección de objetos de configuración. Solo lectura.|1.4|
|[settingCollection](../excel/settingcollection.md)|_Método_ > [add(key: string, value: (any)[])](../excel/settingcollection.md#addkey-string-value-any)|Establece o agrega la configuración especificada en el libro.|1.4|
|[settingCollection](../excel/settingcollection.md)|_Método_ > [getCount()](../excel/settingcollection.md#getcount)|Obtiene el número de configuraciones de una colección.|1.4|
|[settingCollection](../excel/settingcollection.md)|_Método_ > [getItem(key: string)](../excel/settingcollection.md#getitemkey-string)|Obtiene una entrada de configuración mediante la clave.|1.4|
|[settingCollection](../excel/settingcollection.md)|_Método_ > [getItemOrNullObject(key: string)](../excel/settingcollection.md#getitemornullobjectkey-string)|Obtiene una entrada de configuración mediante la clave. Si el valor no existe, devolverá un objeto NULL.|1.4|
|[settingsChangedEventArgs](../excel/settingschangedeventargs.md)|_Relación_ > settings|Obtiene el objeto Setting que representa el enlace que ha generado el evento SettingsChanged.|1.4|
|[tableCollection](../excel/tablecollection.md)|_Método_ > [getCount()](../excel/tablecollection.md#getcount)|Obtiene el número de tablas de la colección.|1.4|
|[tableCollection](../excel/tablecollection.md)|_Método_ > [getItemOrNullObject(key: number or string)](../excel/tablecollection.md#getitemornullobjectkey-number-or-string)|Obtiene una tabla por nombre o identificador. Si la tabla no existe, devolverá un objeto NULL.|1.4|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_Método_ > [getCount()](../excel/tablecolumncollection.md#getcount)|Obtiene el número de columnas de la tabla.|1.4|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_Método_ > [getItemOrNullObject(key: number or string)](../excel/tablecolumncollection.md#getitemornullobjectkey-number-or-string)|Obtiene un objeto de columna por nombre o identificador. Si la columna no existe, devolverá un objeto NULL.|1.4|
|[tableRowCollection](../excel/tablerowcollection.md)|_Método_ > [getCount()](../excel/tablerowcollection.md#getcount)|Obtiene el número de filas de la tabla.|1.4|
|[workbook](../excel/workbook.md)|_Relación_ > settings|Representa una colección de configuraciones asociadas con el libro. Solo lectura.|1.4|
|[worksheet](../excel/worksheet.md)|_Relación_ > names|Colección de nombres en el ámbito de la hoja de cálculo actual. Solo lectura.|1.4|
|[worksheet](../excel/worksheet.md)|_Método_ > [getUsedRangeOrNullObject(valuesOnly: bool)](../excel/worksheet.md#getusedrangeornullobjectvaluesonly-bool)|El rango usado es el rango más pequeño que abarque todas las celdas que tengan asignado un valor o un formato. Si toda la hoja está en blanco, esta función devolverá un objeto NULL.|1.4|
|[worksheetCollection](../excel/worksheetcollection.md)|_Método_ > [getCount(visibleOnly: bool)](../excel/worksheetcollection.md#getcountvisibleonly-bool)|Obtiene el número de hojas de cálculo de la colección.|1.4|
|[worksheetCollection](../excel/worksheetcollection.md)|_Método_ > [getItemOrNullObject(key: string)](../excel/worksheetcollection.md#getitemornullobjectkey-string)|Obtiene un objeto de hoja de cálculo mediante su nombre o identificador. Si la hoja de cálculo no existe, devolverá un objeto NULL.|1.4|



## <a name="whats-new-in-excel-javascript-api-13"></a>Novedades de la API de JavaScript de Excel 1.3
Las siguientes son las nuevas incorporaciones a las API de JavaScript de Excel en el conjunto de requisitos 1.3.

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:----|:----|:----|:----|
|[binding](../excel/binding.md)|_Método_ > [delete()](../excel/binding.md#delete)|Elimina el enlace.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Método_ > [add(range: Range or string, bindingType: string, id: string)](../excel/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|Agregar un enlace nuevo a un intervalo determinado.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Método_ > [addFromNamedItem(name: string, bindingType: string, id: string)](../excel/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|Agregar un enlace nuevo basándose en un elemento con nombre del libro.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Método_ > [addFromSelection(bindingType: string, id: string)](../excel/bindingcollection.md#addfromselectionbindingtype-string-id-string)|Agregar un enlace nuevo basándose en la selección actual.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Método_ > [getItemOrNull(id: string)](../excel/bindingcollection.md#getitemornullid-string)|Obtiene un objeto de enlace por identificador. Si el objeto de enlace no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[chartCollection](../excel/chartcollection.md)|_Método_ > [getItemOrNull(name: string)](../excel/chartcollection.md#getitemornullname-string)|Obtiene un gráfico mediante su nombre. Si hay varias tablas con el mismo nombre, se devolverá la primera.|1.3|
|[namedItemCollection](../excel/nameditemcollection.md)|_Método_ > [getItemOrNull(name: string)](../excel/nameditemcollection.md#getitemornullname-string)|Obtiene un objeto NamedItem mediante su nombre. Si el objeto NamedItem no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[pivotTable](../excel/pivottable.md)|_Propiedad_ > name|Nombre la tabla dinámica.|1.3|
|[pivotTable](../excel/pivottable.md)|_Relación_ > worksheet|La hoja de cálculo que contiene la tabla dinámica actual. Solo lectura.|1.3|
|[pivotTable](../excel/pivottable.md)|_Método_ > [refresh()](../excel/pivottable.md#refresh)|Actualiza la tabla dinámica.|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Propiedad_ > items|Una colección de objetos de tabla dinámica. Solo lectura.|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Método_ > [getItem(name: string)](../excel/pivottablecollection.md#getitemname-string)|Obtiene una tabla dinámica por nombre.|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Método_ > [getItemOrNull(name: string)](../excel/pivottablecollection.md#getitemornullname-string)|Obtiene una tabla dinámica por nombre. Si la tabla dinámica no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[range](../excel/range.md)|_Método_ > [getIntersectionOrNull(anotherRange: Range or string)](../excel/range.md#getintersectionornullanotherrange-range-or-string)|Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados. Si no se encuentra ninguna intersección, se devolverá un objeto NULL.|1.3|
|[range](../excel/range.md)|_Método_ > [getVisibleView()](../excel/range.md#getvisibleview)|Representa las filas visibles del intervalo actual.|1.3|
|[rangeView](../excel/rangeview.md)|_Propiedad_ > cellAddresses|Representa las direcciones de celda de RangeView. Solo lectura.|1.3|
|[rangeView](../excel/rangeview.md)|_Propiedad_ > columnCount|Devuelve el número de columnas visibles. Solo lectura.|1.3|
|[rangeView](../excel/rangeview.md)|_Propiedad_ > formulas|Representa la fórmula en notación de estilo A1.|1.3|
|[rangeView](../excel/rangeview.md)|_Propiedad_ > formulasLocal|Representa la fórmula en notación de estilo A1, en el idioma del usuario y en la configuración regional del formato numérico.  Por ejemplo, la fórmula "=SUM(A1, presentada en 1.5)" en inglés se convertiría en "=SUMME(A1; 1,5)" en alemán.|1.3|
|[rangeView](../excel/rangeview.md)|_Propiedad_ > formulasR1C1|Representa la fórmula en notación de estilo R1C1.|1.3|
|[rangeView](../excel/rangeview.md)|_Propiedad_ > index|Devuelve un valor que representa el índice de RangeView. Solo lectura.|1.3|
|[rangeView](../excel/rangeview.md)|_Propiedad_ > numberFormat|Representa el código de formato numérico de Excel para la celda especificada.|1.3|
|[rangeView](../excel/rangeview.md)|_Propiedad_ > rowCount|Devuelve el número de filas visibles. Solo lectura.|1.3|
|[rangeView](../excel/rangeview.md)|_Propiedad_ > text|Valores de texto del rango especificado. El valor Text no dependerá del ancho de la celda. La sustitución del signo # que tiene lugar en la interfaz de usuario de Excel no afectará al valor de texto devuelto por la API. Solo lectura.|1.3|
|[rangeView](../excel/rangeview.md)|_Propiedad_ > valueTypes|Representa el tipo de datos de cada celda. Solo lectura. Los valores posibles son: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](../excel/rangeview.md)|_Propiedad_ > values|Representa los valores sin formato de la vista del intervalo especificado. Los datos devueltos pueden ser de tipo cadena, número o booleano. La celda que contenga un error devolverá la cadena de error.|1.3|
|[rangeView](../excel/rangeview.md)|_Relación_ > rows|Representa una colección de vistas de intervalo asociadas a este. Solo lectura.|1.3|
|[rangeView](../excel/rangeview.md)|_Método_ > [getRange()](../excel/rangeview.md#getrange)|Obtiene el intervalo primario asociado al RangeView actual.|1.3|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_Propiedad_ > items|Una colección de objetos RangeView. Solo lectura.|1.3|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_Método_ > [getItemAt(index: number)](../excel/rangeviewcollection.md#getitematindex-number)|Obtiene una fila RangeView mediante su índice. Indexado con cero.|1.3|
|[setting](../excel/setting.md)|_Propiedad_ > key|Devuelve una clave que representa el identificador de la configuración. Solo lectura.|1.3|
|[setting](../excel/setting.md)|_Método_ > [delete()](../excel/setting.md#delete)|Elimina la configuración.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Propiedad_ > items|Una colección de objetos de configuración. Solo lectura.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Método_ > [getItem(key: string)](../excel/settingcollection.md#getitemkey-string)|Obtiene una entrada de configuración mediante la clave.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Método_ > [getItemOrNull(key: string)](../excel/settingcollection.md#getitemornullkey-string)|Obtiene una entrada de configuración mediante la clave. Si la configuración no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Método_ > [set(key: string, value: string)](../excel/settingcollection.md#setkey-string-value-string)|Establece o agrega la configuración especificada en el libro.|1.3|
|[settingsChangedEventArgs](../excel/settingschangedeventargs.md)|_Relación_ > settingCollection|Obtiene el objeto Setting que representa el enlace que ha generado el evento SettingsChanged.|1.3|
|[table](../excel/table.md)|_Propiedad_ > highlightFirstColumn|Indica si la primera columna contiene un formato especial.|1.3|
|[table](../excel/table.md)|_Propiedad_ > highlightLastColumn|Indica si la última columna contiene un formato especial.|1.3|
|[table](../excel/table.md)|_Propiedad_ > showBandedColumns|Indica si las columnas muestran un formato con bandas en el que las columnas impares están resaltadas de manera diferente que las pares para facilitar la lectura de la tabla.|1.3|
|[table](../excel/table.md)|_Propiedad_ > showBandedRows|Indica si las filas muestran un formato con bandas en el que las filas impares están resaltadas de manera diferente que las pares para facilitar la lectura de la tabla.|1.3|
|[table](../excel/table.md)|_Propiedad_ > showFilterButton|Indica si los botones de filtro son visibles en la parte superior de cada encabezado de columna. Esta configuración solo se permite si la tabla contiene una fila de encabezado.|1.3|
|[tableCollection](../excel/tablecollection.md)|_Método_ > [getItemOrNull(key: number or string)](../excel/tablecollection.md#getitemornullkey-number-or-string)|Obtiene una tabla por nombre o identificador. Si la tabla no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_Método_ > [getItemOrNull(key: number or string)](../excel/tablecolumncollection.md#getitemornullkey-number-or-string)|Obtiene un objeto de columna por nombre o identificador. Si la columna no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[workbook](../excel/workbook.md)|_Relación_ > pivotTables|Representa una colección de tablas dinámicas asociadas con el libro. Solo lectura.|1.3|
|[workbook](../excel/workbook.md)|_Relación_ > settings|Representa una colección de configuraciones asociadas con el libro. Solo lectura.|1.3|
|[worksheet](../excel/worksheet.md)|_Relación_ > pivotTables|Colección de tablas dinámicas que forman parte de la hoja de cálculo. Solo lectura.|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Novedades de la API de JavaScript de Excel 1.2
Las siguientes son las nuevas incorporaciones a las API de JavaScript de Excel en el conjunto de requisitos 1.2.

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](../excel/chart.md)|_Propiedad_ > id|Obtiene un gráfico en función de su posición en la colección. Solo lectura.|1.2|
|[chart](../excel/chart.md)|_Relación_ > worksheet|La hoja de cálculo que contiene el gráfico actual. Solo lectura.|1.2|
|[chart](../excel/chart.md)|_Método_ > [getImage(height: number, width: number, fittingMode: string)](../excel/chart.md#getimageheight-number-width-number-fittingmode-string)|Representa el gráfico como una imagen con codificación Base64 al escalar el gráfico a las dimensiones especificadas.|1.2|
|[filter](../excel/filter.md)|_Relación_ > criteria|Filtro aplicado actualmente en la columna especificada. Solo lectura.|1.2|
|[filter](../excel/filter.md)|_Método_ > [apply(criteria: FilterCriteria)](../excel/filter.md#applycriteria-filtercriteria)|Aplicar los criterios de filtro especificados en la columna especificada.|1.2|
|[filter](../excel/filter.md)|_Método_ > [applyBottomItemsFilter(count: number)](../excel/filter.md#applybottomitemsfiltercount-number)|Aplicar un filtro de "Elemento inferior" a la columna para el número de elementos especificado.|1.2|
|[filter](../excel/filter.md)|_Método_ > [applyBottomPercentFilter(percent: number)](../excel/filter.md#applybottompercentfilterpercent-number)|Aplicar un filtro de "Porcentaje inferior" a la columna para el porcentaje de elementos especificado.|1.2|
|[filter](../excel/filter.md)|_Método_ > [applyCellColorFilter(color: string)](../excel/filter.md#applycellcolorfiltercolor-string)|Aplicar un filtro de "Color de celda" a la columna para el color especificado.|1.2|
|[filter](../excel/filter.md)|_Método_ > [applyCustomFilter(criteria1: string, criteria2: string, oper: string)](../excel/filter.md#applycustomfiltercriteria1-string-criteria2-string-oper-string)|Aplicar un filtro de "Icono" a la columna para las cadenas de criterios especificadas.|1.2|
|[filter](../excel/filter.md)|_Método_ > [applyDynamicFilter(criteria: string)](../excel/filter.md#applydynamicfiltercriteria-string)|Aplicar un filtro "Dinámico" a la columna.|1.2|
|[filter](../excel/filter.md)|_Método_ > [applyFontColorFilter(color: string)](../excel/filter.md#applyfontcolorfiltercolor-string)|Aplicar un filtro de "Color de fuente" a la columna para el color especificado.|1.2|
|[filter](../excel/filter.md)|_Método_ > [applyIconFilter(icon: Icon)](../excel/filter.md#applyiconfiltericon-icon)|Aplicar un filtro de "Icono" a la columna para el icono especificado.|1.2|
|[filter](../excel/filter.md)|_Método_ > [applyTopItemsFilter(count: number)](../excel/filter.md#applytopitemsfiltercount-number)|Aplicar un filtro de "Elemento superior" a la columna para el número de elementos especificado.|1.2|
|[filter](../excel/filter.md)|_Método_ > [applyTopPercentFilter(percent: number)](../excel/filter.md#applytoppercentfilterpercent-number)|Aplicar un filtro de "Porcentaje superior" a la columna para el porcentaje de elementos especificado.|1.2|
|[filter](../excel/filter.md)|_Método_ > [applyValuesFilter(values: ()[])](../excel/filter.md#applyvaluesfiltervalues-)|Aplicar un filtro de "Valores" a la columna para los valores especificados.|1.2|
|[filter](../excel/filter.md)|_Método_ > [clear()](../excel/filter.md#clear)|Borrar el filtro de la columna especificada.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propiedad_ > color|Cadena de color HTML que se usa para filtrar las celdas. Se usa con el filtrado de "cellColor" y "fontColor".|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propiedad_ > criterion1|Primer criterio usado para filtrar los datos. Se usa como un operador en el caso del filtrado "personalizado".|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propiedad_ > criterion2|Segundo criterio usado para filtrar los datos. Solo se usa como un operador en el caso del filtrado "personalizado".|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propiedad_ > dynamicCriteria|Criterios dinámicos del conjunto Excel.DynamicFilterCriteria que se van a aplicar a esta columna. Se usa con el filtrado "dinámico". Los valores posibles son: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propiedad_ > filterOn|Propiedad usada por el filtro para determinar si los valores deben permanecer visibles. Los valores posibles son: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propiedad_ > operator|Operador usado para combinar el criterio 1 y 2 cuando se usa el filtrado "personalizado". Los valores posibles son: And, Or.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propiedad_ > values|Conjunto de valores que se van a usar como parte del filtrado de "valores".|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Relación_ > icon|Icono usado para filtrar las celdas. Se usa con el filtrado de "icono".|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_Propiedad_ > date|La fecha en formato ISO8601 usada para filtrar los datos.|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_Propiedad_ > specificity|El grado de especificidad de la fecha que se usará para mantener datos. Por ejemplo, si la fecha es 02-04-2005 y la especificidad se establece en "mes", la operación de filtrado conservará todas las filas con fecha de abril de 2005. Los valores posibles son: Year, Monday, Day, Hour, Minute, Second.|1.2|
|[formatProtection](../excel/formatprotection.md)|_Propiedad_ > formulaHidden|Indica si Excel oculta la fórmula de las celdas del rango. Un valor null indica que el rango no tiene una configuración de fórmula oculta uniforme.|1.2|
|[formatProtection](../excel/formatprotection.md)|_Propiedad_ > locked|Indica si Excel bloquea las celdas del objeto. Un valor nulo indica que todo el rango no tiene una configuración de bloqueo uniforme.|1.2|
|[icon](../excel/icon.md)|_Propiedad_ > index|Representa el índice del icono en el conjunto concreto.|1.2|
|[icon](../excel/icon.md)|_Propiedad_ > set|Representa el conjunto al que pertenece el icono. Los valores posibles son: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](../excel/range.md)|_Propiedad_ > columnHidden|Representa si todas las columnas del intervalo actual están ocultas.|1.2|
|[range](../excel/range.md)|_Propiedad_ > formulasR1C1|Representa la fórmula en notación de estilo R1C1.|1.2|
|[range](../excel/range.md)|_Propiedad_ > hidden|Representa si todas las celdas del rango actual están ocultas. Solo lectura.|1.2|
|[range](../excel/range.md)|_Propiedad_ > rowHidden|Representa si todas las filas del intervalo actual están ocultas.|1.2|
|[range](../excel/range.md)|_Relación_ > sort|Representa la ordenación del intervalo del intervalo actual. Solo lectura.|1.2|
|[range](../excel/range.md)|_Método_ > [merge(across: bool)](../excel/range.md#mergeacross-bool)|Combina las celdas del intervalo en una región de la hoja de cálculo.|1.2|
|[range](../excel/range.md)|_Método_ > [unmerge()](../excel/range.md#unmerge)|Separa las celdas del intervalo en celdas independientes.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Propiedad_ > columnWidth|Obtiene o establece el ancho de todas las columnas del rango. Si los anchos de columna no son uniformes, se devolverá null.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Propiedad_ > rowHeight|Obtiene o establece el alto de todas las filas del rango. Si los altos de fila no son uniformes, se devolverá null.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Relación_ > protection|Devuelve el objeto de protección de formato de un rango. Solo lectura.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Método_ > [autofitColumns()](../excel/rangeformat.md#autofitcolumns)|Cambia el ancho de las columnas del intervalo actual para obtener el ajuste perfecto (según los datos actuales de las columnas).|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Método_ > [autofitRows()](../excel/rangeformat.md#autofitrows)|Cambia el alto de las filas del intervalo actual para obtener el ajuste perfecto (según los datos actuales de las columnas).|1.2|
|[rangeReference](../excel/rangereference.md)|_Propiedad_ > address|Representa las filas visibles del intervalo actual.|1.2|
|[rangeSort](../excel/rangesort.md)|_Método_ > [apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](../excel/rangesort.md#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|Realiza una operación de ordenación.|1.2|
|[sortField](../excel/sortfield.md)|_Propiedad_ > ascending|Representa si la ordenación se realiza en orden ascendente.|1.2|
|[sortField](../excel/sortfield.md)|_Propiedad_ > color|Representa el color que es el destino de la condición si la ordenación se realiza según la fuente o el color de celda.|1.2|
|[sortField](../excel/sortfield.md)|_Propiedad_ > dataOption|Representa opciones de ordenación adicionales para este campo. Los valores posibles son: Normal, TextAsNumber.|1.2|
|[sortField](../excel/sortfield.md)|_Propiedad_ > key|Representa la columna (o fila, según la orientación de ordenación) en que se encuentra la condición. Se representa como un desplazamiento de la primera columna (o fila).|1.2|
|[sortField](../excel/sortfield.md)|_Propiedad_ > sortOn|Representa el tipo de ordenación de esta condición. Los valores posibles son: Value, CellColor, FontColor, Icon.|1.2|
|[sortField](../excel/sortfield.md)|_Relación_ > icon|Representa el icono que es el destino de la condición si la ordenación se realiza según el icono de la celda.|1.2|
|[table](../excel/table.md)|_Relación_ > sort|Representa la ordenación de la tabla. Solo lectura.|1.2|
|[table](../excel/table.md)|_Relación_ > worksheet|Hoja de cálculo que contiene la tabla actual. Solo lectura.|1.2|
|[table](../excel/table.md)|_Método_ > [clearFilters()](../excel/table.md#clearfilters)|Borra todos los filtros aplicados actualmente en la tabla.|1.2|
|[table](../excel/table.md)|_Método_ > [convertToRange()](../excel/table.md#converttorange)|Convierte la tabla en un rango de celdas normal. Se conservan todos los datos.|1.2|
|[table](../excel/table.md)|_Método_ > [reapplyFilters()](../excel/table.md#reapplyfilters)|Vuelve a aplicar todos los filtros aplicados actualmente en la tabla.|1.2|
|[tableColumn](../excel/tablecolumn.md)|_Relación_ > filter|Recupera el filtro aplicado a la columna. Solo lectura.|1.2|
|[tableSort](../excel/tablesort.md)|_Propiedad_ > matchCase|Indica si última ordenación de la tabla distinguía mayúsculas de minúsculas. Solo lectura.|1.2|
|[tableSort](../excel/tablesort.md)|_Propiedad_ > method|Representa el método de ordenación de caracteres chinos usado por última vez para ordenar la tabla. Solo lectura. Los valores posibles son: PinYin, StrokeCount.|1.2|
|[tableSort](../excel/tablesort.md)|_Relación_ > fields|Representa las condiciones actuales que se usaron por última vez para ordenar la tabla. Solo lectura.|1.2|
|[tableSort](../excel/tablesort.md)|_Método_ > [apply(fields: SortField[], matchCase: bool, method: string)](../excel/tablesort.md#applyfields-sortfield-matchcase-bool-method-string)|Realiza una operación de ordenación.|1.2|
|[tableSort](../excel/tablesort.md)|_Método_ > [clear()](../excel/tablesort.md#clear)|Borra la ordenación que se aplica actualmente en la tabla. Aunque esto no modifica la ordenación de la tabla, borra el estado de los botones de encabezado.|1.2|
|[tableSort](../excel/tablesort.md)|_Método_ > [reapply()](../excel/tablesort.md#reapply)|Vuelve a aplicar los parámetros de ordenación actuales a la tabla.|1.2|
|[workbook](../excel/workbook.md)|_Relación_ > functions|Representa una instancia de aplicación de Excel que contiene este libro. Solo lectura.|1.2|
|[worksheet](../excel/worksheet.md)|_Relación_ > protection|Devuelve el objeto de protección de hoja de una hoja de cálculo. Solo lectura.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Propiedad_ > protected|Indica si la hoja de cálculo está protegida. Solo lectura. Solo lectura.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Relación_ > options|Opciones de protección de la hoja. Solo lectura.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Método_ > [protect(options: WorksheetProtectionOptions)](../excel/worksheetprotection.md#protectoptions-worksheetprotectionoptions)|Protege una hoja de cálculo. Produce un error si se ha protegido la hoja de cálculo.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Método_ > [unprotect()](../excel/worksheetprotection.md#unprotect)|Desprotege una hoja de cálculo.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propiedad_ > allowAutoFilter|Representa la opción de protección de la hoja de cálculo que permite usar la característica de filtro automático.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propiedad_ > allowDeleteColumns|Representa la opción de protección de la hoja de cálculo que permite eliminar columnas.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propiedad_ > allowDeleteRows|Representa la opción de protección de la hoja de cálculo que permite eliminar filas.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propiedad_ > allowFormatCells|Representa la opción de protección de la hoja de cálculo que permite aplicar formato a celdas.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propiedad_ > allowFormatColumns|Representa la opción de protección de la hoja de cálculo que permite aplicar formato a columnas.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propiedad_ > allowFormatRows|Representa la opción de protección de la hoja de cálculo que permite aplicar formato a filas.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propiedad_ > allowInsertColumns|Representa la opción de protección de la hoja de cálculo que permite insertar columnas.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propiedad_ > allowInsertHyperlinks|Representa la opción de protección de la hoja de cálculo que permite insertar hipervínculos.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propiedad_ > allowInsertRows|Representa la opción de protección de la hoja de cálculo que permite insertar filas.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propiedad_ > allowPivotTables|Representa la opción de protección de la hoja de cálculo que permite usar la característica de tabla dinámica.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propiedad_ > allowSort|Representa la opción de protección de la hoja de cálculo que permite usar la característica de ordenación.|1.2|

## <a name="excel-javascript-api-11"></a>API de JavaScript de Excel 1.1
API de JavaScript de Excel 1.1 es la primera versión de la API. Para obtener más información sobre la API, consulte los temas de referencia de API de JavaScript de Excel.  

## <a name="additional-resources"></a>Recursos adicionales

- [Especificar los hosts de Office y los requisitos de la API](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Manifiesto XML de complementos para Office](../../docs/overview/add-in-manifests.md)
