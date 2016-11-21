# <a name="excel-javascript-api-requirement-sets"></a>Conjuntos de requisitos de la API de JavaScript de Excel

Los conjuntos de requisitos son grupos de miembros de la API con nombre. Los complementos de Office usan los conjuntos de requisitos especificados en el manifiesto o usan una comprobación en tiempo de ejecución para determinar si un host de Office admite las API necesarias para el complemento. Para obtener más información, consulte [Specify Office hosts and API requirements (Especificar hosts de Office y requisitos de la API)](../docs/overview/specify-office-hosts-and-api-requirements.md).

Los complementos de Excel se ejecutan en varias versiones de Office, incluida Office 2016 para Windows, Office para iPad, Office para Mac y Office Online. En la siguiente tabla se enumeran los conjuntos de requisitos de Excel, las aplicaciones de host de Office que admiten ese conjunto de requisitos y el número o las versiones de compilación de esas aplicaciones. 

|  Conjunto de requisitos  |  Office 2016 para Windows*  |  Office 2016 para iPad  |  Office 2016 para Mac  | Office Online  |
|:-----|-----|:-----|:-----|:-----|
| ExcelApi 1.3  | Versión 1608 o posterior| 1.27 o posterior |  15.27 o posterior| Septiembre de 2016 | 
| ExcelApi 1.2  | Versión 1601 o posterior | 1.21 o posterior | 15.22 o posterior| Enero de 2016 |
| ExcelApi 1.1  | Versión 1509 (compilación 4266.1001) o posterior | 1.19 o posterior | 15.20 o posterior| Enero de 2016 |

> &#42; **Nota**: El número de compilación para Office 2016 que se ha instalado mediante MSI es 16.0.4266.1001. Esta versión solo contiene el conjunto de requisitos de ExcelApi 1.1.

Para obtener más información sobre las versiones y números de compilación, consulte:

- [Números de versión y compilación de las versiones del canal de actualización para los clientes de Office 365](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [¿Qué versión de Office estoy usando?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Dónde puede encontrar el número de versión y de compilación de una aplicación de cliente de Office 365](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API de Office
Para obtener información sobre los conjuntos de requisitos comunes de la API, consulte [Office common API requirement sets (Conjuntos de requisitos comunes de la API de Office)](office-add-in-requirement-sets.md).

## <a name="whats-new-in-excel-javascript-api-13"></a>Novedades de la API de JavaScript de Excel 1.3 
Las siguientes son las nuevas incorporaciones a las API de JavaScript de Excel en el conjunto de requisitos 1.3. 

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:----|:----|:----|:----|
|[binding](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/binding.md)|_Método_ > [delete()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/binding.md#delete)|Elimina el enlace.|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_Método_ > [add(range: Range or string, bindingType: string, id: string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|Agregar un enlace nuevo a un intervalo determinado.|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_Método_ > [addFromNamedItem(name: string, bindingType: string, id: string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|Agregar un enlace nuevo basándose en un elemento con nombre del libro.|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_Método_ > [addFromSelection(bindingType: string, id: string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#addfromselectionbindingtype-string-id-string)|Agregar un enlace nuevo basándose en la selección actual.|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_Método_ > [getItemOrNull(id: string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#getitemornullid-string)|Obtiene un objeto de enlace por identificador. Si el objeto de enlace no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[chartCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/chartcollection.md)|_Método_ > [getItemOrNull(name: string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/chartcollection.md#getitemornullname-string)|Obtiene un gráfico mediante su nombre. Si hay varios gráficos con el mismo nombre, se devolverá el primero.|1.3|
|[namedItemCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/nameditemcollection.md)|_Método_ > [getItemOrNull(name: string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/nameditemcollection.md#getitemornullname-string)|Obtiene un objeto NamedItem mediante su nombre. Si el objeto NamedItem no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[pivotTable](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md)|_Propiedad_ > name|Nombre la tabla dinámica.|1.3|
|[pivotTable](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md)|_Relación_ > worksheet|La hoja de cálculo que contiene la tabla dinámica actual. Solo lectura.|1.3|
|[pivotTable](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md)|_Método_ > [refresh()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md#refresh)|Actualiza la tabla dinámica.|1.3|
|[pivotTableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md)|_Propiedad_ > items|Una colección de objetos de tabla dinámica. Solo lectura.|1.3|
|[pivotTableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md)|_Método_ > [getItem(name: string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md#getitemname-string)|Obtiene una tabla dinámica por nombre.|1.3|
|[pivotTableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md)|_Método_ > [getItemOrNull(name: string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md#getitemornullname-string)|Obtiene una tabla dinámica por nombre. Si la tabla dinámica no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[range](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md)|_Método_ > [getIntersectionOrNull(anotherRange: Range or string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md#getintersectionornullanotherrange-range-or-string)|Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados. Si no se encuentra ninguna intersección, se devolverá un objeto NULL.|1.3|
|[range](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md)|_Método_ > [getVisibleView()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md#getvisibleview)|Representa las filas visibles del intervalo actual.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propiedad_ > cellAddresses|Representa las direcciones de celda de RangeView. Solo lectura.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propiedad_ > columnCount|Devuelve el número de columnas visibles. Solo lectura.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propiedad_ > formulas|Representa la fórmula en notación de estilo A1.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propiedad_ > formulasLocal|Representa la fórmula en notación de estilo A1, en el idioma del usuario y en la configuración regional del formato numérico.  Por ejemplo, la fórmula "=SUM(A1, presentada en 1.5)" en inglés se convertiría en "=SUMME(A1; 1,5)" en alemán.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propiedad_ > formulasR1C1|Representa la fórmula en notación de estilo R1C1.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propiedad_ > index|Devuelve un valor que representa el índice de RangeView. Solo lectura.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propiedad_ > numberFormat|Representa el código de formato numérico de Excel para la celda especificada.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propiedad_ > rowCount|Devuelve el número de filas visibles. Solo lectura.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propiedad_ > text|Valores de texto del rango especificado. El valor Text no dependerá del ancho de la celda. La sustitución del signo # que tiene lugar en la interfaz de usuario de Excel no afectará al valor de texto devuelto por la API. Solo lectura.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propiedad_ > valueTypes|Representa el tipo de datos de cada celda. Solo lectura. Los valores posibles son: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propiedad_ > values|Representa los valores sin formato de la vista del intervalo especificado. Los datos devueltos pueden ser de tipo cadena, número o booleano. La celda que contenga un error devolverá la cadena de error.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Relación_ > rows|Representa una colección de vistas de intervalo asociadas a este. Solo lectura.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Método_ > [getRange()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md#getrange)|Obtiene el intervalo primario asociado al RangeView actual.|1.3|
|[rangeViewCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeviewcollection.md)|_Propiedad_ > items|Una colección de objetos RangeView. Solo lectura.|1.3|
|[rangeViewCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeviewcollection.md)|_Método_ > [getItemAt(index: number)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeviewcollection.md#getitematindex-number)|Obtiene una fila RangeView mediante su índice. Indexado con cero.|1.3|
|[setting](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/setting.md)|_Propiedad_ > key|Devuelve una clave que representa el identificador de la configuración. Solo lectura.|1.3|
|[setting](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/setting.md)|_Método_ > [delete()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/setting.md#delete)|Elimina la configuración.|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_Propiedad_ > items|Una colección de objetos de configuración. Solo lectura.|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_Método_ > [getItem(key: string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md#getitemkey-string)|Obtiene una entrada de configuración mediante la clave.|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_Método_ > [getItemOrNull(key: string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md#getitemornullkey-string)|Obtiene una entrada de configuración mediante la clave. Si la configuración no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_Método_ > [set(key: string, value: string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md#setkey-string-value-string)|Establece o agrega la configuración especificada en el libro.|1.3|
|[settingsChangedEventArgs](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingschangedeventargs.md)|_Relación_ > settingCollection|Obtiene el objeto Setting que representa el enlace que ha generado el evento SettingsChanged.|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_Propiedad_ > highlightFirstColumn|Indica si la primera columna contiene un formato especial.|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_Propiedad_ > highlightLastColumn|Indica si la última columna contiene un formato especial.|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_Propiedad_ > showBandedColumns|Indica si las columnas muestran un formato con bandas en el que las columnas impares están resaltadas de manera diferente que las pares para facilitar la lectura de la tabla.|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_Propiedad_ > showBandedRows|Indica si las filas muestran un formato con bandas en el que las filas impares están resaltadas de manera diferente que las pares para facilitar la lectura de la tabla.|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_Propiedad_ > showFilterButton|Indica si los botones de filtro son visibles en la parte superior de cada encabezado de columna. Esta configuración solo se permite si la tabla contiene una fila de encabezado.|1.3|
|[tableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecollection.md)|_Método_ > [getItemOrNull(key: number or string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecollection.md#getitemornullkey-number-or-string)|Obtiene una tabla por nombre o identificador. Si la tabla no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[tableColumnCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecolumncollection.md)|_Método_ > [getItemOrNull(key: number or string)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecolumncollection.md#getitemornullkey-number-or-string)|Obtiene un objeto de columna por nombre o identificador. Si la columna no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[workbook](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/workbook.md)|_Relación_ > pivotTables|Representa una colección de tablas dinámicas asociadas con el libro. Solo lectura.|1.3|
|[workbook](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/workbook.md)|_Relación_ > settings|Representa una colección de configuraciones asociadas con el libro. Solo lectura.|1.3|
|[worksheet](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/worksheet.md)|_Relación_ > pivotTables|Colección de tablas dinámicas que forman parte de la hoja de cálculo. Solo lectura.|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Novedades de la API de JavaScript de Excel 1.2
Las siguientes son las nuevas incorporaciones a las API de JavaScript de Excel en el conjunto de requisitos 1.2. 

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](../excel/chart.md)|_Propiedad_ > id|Obtiene un gráfico en función de su posición en la colección. Solo lectura.|1.2|
|[chart](../excel/chart.md)|_Relación_ > worksheet|La hoja de cálculo que contiene el gráfico actual. Solo lectura.|1.2|
|[chart](../excel/chart.md)|_Método_ > [getImage(height: number, width: number, fittingMode: string)](../excel/chart.md#getimageheight-number-width-number-fittingmode-string)|Representa el gráfico como una imagen con codificación Base64 al escalar el gráfico a las dimensiones especificadas.|1.2|
|[filter](../excel/filter.md)|_Relación_ > criteria|El filtro aplicado actualmente en la columna especificada. Solo lectura.|1.2|
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
|[formatProtection](../excel/formatprotection.md)|_Propiedad_ > locked|Indica si Excel bloquea las celdas del objeto. Un valor NULL indica que el intervalo no tiene una configuración de bloqueo uniforme.|1.2|
|[icon](../excel/icon.md)|_Propiedad_ > index|Representa el índice del icono en el conjunto concreto.|1.2|
|[icon](../excel/icon.md)|_Propiedad_ > set|Representa el conjunto al que pertenece el icono. Los valores posibles son: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](../excel/range.md)|_Propiedad_ > columnHidden|Representa si todas las columnas del intervalo actual están ocultas.|1.2|
|[range](../excel/range.md)|_Propiedad_ > formulasR1C1|Representa la fórmula en notación de estilo R1C1.|1.2|
|[range](../excel/range.md)|_Propiedad_ > hidden|Representa si todas las celdas del intervalo actual están ocultas. Solo lectura.|1.2|
|[range](../excel/range.md)|_Propiedad_ > rowHidden|Representa si todas las filas del intervalo actual están ocultas.|1.2|
|[range](../excel/range.md)|_Relación_ > sort|Representa la ordenación del intervalo del intervalo actual. Solo lectura.|1.2|
|[range](../excel/range.md)|_Método_ > [merge(across: bool)](../excel/range.md#mergeacross-bool)|Combina las celdas del intervalo en una región de la hoja de cálculo.|1.2|
|[range](../excel/range.md)|_Método_ > [unmerge()](../excel/range.md#unmerge)|Separa las celdas del intervalo en celdas independientes.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Propiedad_ > columnWidth|Obtiene o establece el ancho de todas las columnas del intervalo. Si los anchos de columna no son uniformes, se devolverá NULL.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Propiedad_ > rowHeight|Obtiene o establece el alto de todas las filas del intervalo. Si los altos de fila no son uniformes, se devolverá NULL.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Relación_ > protection|Devuelve el objeto de protección de formato de un intervalo. Solo lectura.|1.2|
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
|[table](../excel/table.md)|_Relación_ > worksheet|La hoja de cálculo que contiene la tabla actual. Solo lectura.|1.2|
|[table](../excel/table.md)|_Método_ > [clearFilters()](../excel/table.md#clearfilters)|Borra todos los filtros aplicados actualmente en la tabla.|1.2|
|[table](../excel/table.md)|_Método_ > [convertToRange()](../excel/table.md#converttorange)|Convierte la tabla en un intervalo de celdas normal. Se conservan todos los datos.|1.2|
|[table](../excel/table.md)|_Método_ > [reapplyFilters()](../excel/table.md#reapplyfilters)|Vuelve a aplicar todos los filtros aplicados actualmente en la tabla.|1.2|
|[tableColumn](../excel/tablecolumn.md)|_Relación_ > filter|Recupera el filtro aplicado a la columna. Solo lectura.|1.2|
|[tableSort](../excel/tablesort.md)|_Propiedad_ > matchCase|Indica si la última ordenación de la tabla distinguía mayúsculas de minúsculas. Solo lectura.|1.2|
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

- [Especificar los hosts de Office y los requisitos de la API](../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Manifiesto XML de complementos para Office](https://dev.office.com/docs/add-ins/overview/add-in-manifests)