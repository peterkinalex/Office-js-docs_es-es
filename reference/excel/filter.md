# <a name="filter-object-(javascript-api-for-excel)"></a>Objeto Filter (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS, Office 2016_

Administra el filtrado de la columna de una tabla.

## <a name="properties"></a>Properties

Ninguno

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|criteria|[FilterCriteria](filtercriteria.md)|Filtro aplicado actualmente en la columna especificada. Solo lectura.|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[apply(criteria: FilterCriteria)](#applycriteria-filtercriteria)|void|Aplica los criterios de filtro especificados en la columna especificada. Se puede conseguir la misma funcionalidad con cualquiera de los siguientes métodos auxiliares.|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|Aplica un filtro de "Elemento inferior" a la columna para el número de elementos especificado.|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|Aplica un filtro de "Porcentaje inferior" a la columna para el porcentaje de elementos especificado.|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|Aplica un filtro de "Color de celda" a la columna para el color especificado.|
|[applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)](#applycustomfiltercriteria1-string-criteria2-string-oper-filteroperator)|void|Aplica un filtro de "Icono" a la columna para las cadenas de criterios especificadas.|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|Aplica un filtro "Dinámico" a la columna.|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|Aplica un filtro de "Color de fuente" a la columna para el color especificado.|
|[applyIconFilter(icon: Icon)](#applyiconfiltericon-icon)|void|Aplica un filtro de "Icono" a la columna para el icono especificado.|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|Aplica un filtro de "Elemento superior" a la columna para el número de elementos especificado.|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|Aplica un filtro de "Porcentaje superior" a la columna para el porcentaje de elementos especificado.|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|Aplica un filtro de "Valores" a la columna para los valores especificados.|
|[clear()](#clear)|void|Borra el filtro de la columna especificada.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## <a name="method-details"></a>Detalles del método


### <a name="apply(criteria:-filtercriteria)"></a>apply(criteria: FilterCriteria)
Aplica los criterios de filtro especificados en la columna especificada. Se puede conseguir la misma funcionalidad con cualquiera de los siguientes métodos auxiliares. 

#### <a name="syntax"></a>Sintaxis
```js
filterObject.apply(criteria);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|criterios|FilterCriteria|Criterios que se aplicarán.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="example"></a>Ejemplo
En el ejemplo siguiente se muestra cómo aplicar un filtro personalizado con el método genérico apply().

```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    var filterCriteria = { 
        filterOn: Excel.FilterOn.custom,
        criterion1: ">50",
        operator: Excel.FilterOperator.and,
        criterion2: "<100"
        } 
    column.filter.apply(filterCriteria);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="applybottomitemsfilter(count:-number)"></a>applyBottomItemsFilter(count: number)
Aplica un filtro de "Elemento inferior" a la columna para el número de elementos especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyBottomItemsFilter(count);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|count|number|Número de elementos desde la parte inferior que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="example"></a>Ejemplo
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="applybottompercentfilter(percent:-number)"></a>applyBottomPercentFilter(percent: number)
Aplica un filtro de "Porcentaje inferior" a la columna para el porcentaje de elementos especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyBottomPercentFilter(percent);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|signo de porcentaje|number|Porcentaje de elementos desde la parte inferior que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="example"></a>Ejemplo
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="applycellcolorfilter(color:-string)"></a>applyCellColorFilter(color: string)
Aplica un filtro de "Color de celda" a la columna para el color especificado.


#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyCellColorFilter(color);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|color|string|Color de fondo de las celdas que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="example"></a>Ejemplo
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCellColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="applycustomfilter(criteria1:-string,-criteria2:-string,-oper:-filteroperator)"></a>applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)
Aplica un filtro de "Icono" a la columna para las cadenas de criterios especificadas.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|criteria1|string|Primera cadena de criterios.|
|criteria2|string|Opcional. Segunda cadena de criterios.|
|oper|FilterOperator|Opcional. Operador que describe cómo se combinan los dos criterios.|

#### <a name="returns"></a>Valores devueltos
void


#### <a name="example"></a>Ejemplo
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCustomFilter('>50','<100','and');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="applydynamicfilter(criteria:-string)"></a>applyDynamicFilter(criteria: string)
Aplica un filtro "Dinámico" a la columna.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyDynamicFilter(criteria);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|criterios|string|Criterios dinámicos que se aplicarán.  Los valores posibles son: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate y Yesterday.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="example"></a>Ejemplo
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyDynamicFilter(Excel.DynamicFilterCriteria.aboveAverage);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="applyfontcolorfilter(color:-string)"></a>applyFontColorFilter(color: string)
Aplica un filtro de "Color de fuente" a la columna para el color especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyFontColorFilter(color);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|color|string|Color de fuente de las celdas que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="example"></a>Ejemplo
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyFontColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="applyiconfilter(icon:-icon)"></a>applyIconFilter(icon: Icon)
Aplica un filtro de "Icono" a la columna para el icono especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyIconFilter(icon);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|icono|Icono|Iconos de las celdas que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="example"></a>Ejemplo
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyIconFilter(Excel.icons.fiveArrows.yellowDownInclineArrow);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="applytopitemsfilter(count:-number)"></a>applyTopItemsFilter(count: number)
Aplica un filtro de "Elemento superior" a la columna para el número de elementos especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyTopItemsFilter(count);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|count|number|Número de elementos desde la parte superior que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="example"></a>Ejemplo
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="applytoppercentfilter(percent:-number)"></a>applyTopPercentFilter(percent: number)
Aplica un filtro de "Porcentaje superior" a la columna para el porcentaje de elementos especificado.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyTopPercentFilter(percent);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|signo de porcentaje|number|Porcentaje de elementos desde la parte superior que se van a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="example"></a>Ejemplo
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="applyvaluesfilter(values:-()[])"></a>applyValuesFilter(values: ()[])
Aplica un filtro de "Valores" a la columna para los valores especificados.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.applyValuesFilter(values);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|values|()[]|Lista de valores que se va a mostrar.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="example"></a>Ejemplo
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyValuesFilter(['a','b']);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="clear()"></a>clear()
Desactiva el filtro de la columna especificada.

#### <a name="syntax"></a>Sintaxis
```js
filterObject.clear();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="example"></a>Ejemplo
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="load(param:-object)"></a>load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
object.load(param);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
void
