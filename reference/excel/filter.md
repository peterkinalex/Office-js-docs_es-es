# <a name="filter-object-javascript-api-for-excel"></a>Objeto Filter (API de JavaScript para Excel)

Administra el filtrado de la columna de una tabla.

## <a name="properties"></a>Properties

Ninguno

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|criteria|[FilterCriteria](filtercriteria.md)|Filtro aplicado actualmente en la columna especificada. Solo lectura.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[apply(criteria: FilterCriteria)](#applycriteria-filtercriteria)|void|Aplicar los criterios de filtro especificados en la columna especificada.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|Aplicar un filtro de "Elemento inferior" a la columna para el número de elementos especificado.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|Aplicar un filtro de "Porcentaje inferior" a la columna para el porcentaje de elementos especificado.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|Aplicar un filtro de "Color de celda" a la columna para el color especificado.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyCustomFilter(criteria1: string, criteria2: string, oper: string)](#applycustomfiltercriteria1-string-criteria2-string-oper-string)|void|Aplicar un filtro de "Icono" a la columna para las cadenas de criterios especificadas.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|Aplicar un filtro "Dinámico" a la columna.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|Aplicar un filtro de "Color de fuente" a la columna para el color especificado.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyIconFilter(icon: Icon)](#applyiconfiltericon-icon)|void|Aplicar un filtro de "Icono" a la columna para el icono especificado.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|Aplicar un filtro de "Elemento superior" a la columna para el número de elementos especificado.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|Aplicar un filtro de "Porcentaje superior" a la columna para el porcentaje de elementos especificado.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|Aplicar un filtro de "Valores" a la columna para los valores especificados.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[clear()](#clear)|void|Borrar el filtro de la columna especificada.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy que se ha creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="applycriteria-filtercriteria"></a>apply(criteria: FilterCriteria)
Aplicar los criterios de filtro especificados en la columna especificada.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.apply(criteria);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|criterios|FilterCriteria|Criterios que se aplicarán.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="applybottomitemsfiltercount-number"></a>applyBottomItemsFilter(count: number)
Aplica un filtro de "Elemento inferior" a la columna para el número de elementos especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyBottomItemsFilter(count);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|count|number|Número de elementos desde la parte inferior que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="applybottompercentfilterpercent-number"></a>applyBottomPercentFilter(percent: number)
Aplica un filtro de "Porcentaje inferior" a la columna para el porcentaje de elementos especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyBottomPercentFilter(percent);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|signo de porcentaje|number|Porcentaje de elementos desde la parte inferior que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="applycellcolorfiltercolor-string"></a>applyCellColorFilter(color: string)
Aplica un filtro de "Color de celda" a la columna para el color especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyCellColorFilter(color);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|color|string|Color de fondo de las celdas que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="applycustomfiltercriteria1-string-criteria2-string-oper-string"></a>applyCustomFilter(criteria1: string, criteria2: string, oper: string)
Aplicar un filtro de "Icono" a la columna para las cadenas de criterios especificadas.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|criteria1|string|Primera cadena de criterios.|
|criteria2|string|Opcional. Segunda cadena de criterios.|
|oper|string|Opcional. El operador que describe cómo se combinan los dos criterios.  Los valores posibles son: And, Or.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="applydynamicfiltercriteria-string"></a>applyDynamicFilter(criteria: string)
Aplica un filtro "Dinámico" a la columna.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyDynamicFilter(criteria);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|criterios|string|Criterios dinámicos que se aplicarán.  Los valores posibles son: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate y Yesterday.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="applyfontcolorfiltercolor-string"></a>applyFontColorFilter(color: string)
Aplica un filtro de "Color de fuente" a la columna para el color especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyFontColorFilter(color);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|color|string|Color de fuente de las celdas que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="applyiconfiltericon-icon"></a>applyIconFilter(icon: Icon)
Aplica un filtro de "Icono" a la columna para el icono especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyIconFilter(icon);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|icono|Icono|Iconos de las celdas que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="applytopitemsfiltercount-number"></a>applyTopItemsFilter(count: number)
Aplica un filtro de "Elemento superior" a la columna para el número de elementos especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyTopItemsFilter(count);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|count|number|Número de elementos desde la parte superior que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="applytoppercentfilterpercent-number"></a>applyTopPercentFilter(percent: number)
Aplica un filtro de "Porcentaje superior" a la columna para el porcentaje de elementos especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyTopPercentFilter(percent);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|signo de porcentaje|number|Porcentaje de elementos desde la parte superior que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="applyvaluesfiltervalues-"></a>applyValuesFilter(values: ()[])
Aplica un filtro de "Valores" a la columna para los valores especificados.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyValuesFilter(values);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|values|()[]|Lista de valores que se va a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="clear"></a>clear()
Desactiva el filtro de la columna especificada.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.clear();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

### <a name="loadparam-object"></a>load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
object.load(param);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
void
