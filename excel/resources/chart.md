# Objeto Chart (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Office 2016_

Representa un objeto de gráfico de una hoja de cálculo.

## Propiedades

| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|height|Double|Representa el alto, en puntos, del objeto de gráfico.|
|left|Double|Distancia, en puntos, desde el lado izquierdo del gráfico hasta el origen de la hoja de cálculo.|
|name|string|Representa el nombre de un objeto de gráfico.|
|top|Double|Representa la distancia, en puntos, desde el borde superior del objeto hasta la parte superior de la fila 1 (en una hoja de cálculo) o la parte superior del área del gráfico (en un gráfico).|
|width|double|Representa el ancho, en puntos, del objeto de gráfico.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
| Relación | Tipo|Descripción|
|:---------------|:--------|:----------|
|axes|[ChartAxes](chartaxes.md)|Representa los ejes del gráfico. Solo lectura.|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|Representa la clase DataLabels del gráfico. Solo lectura.|
|format|[ChartAreaFormat](chartareaformat.md)|Encapsula las propiedades de formato del área del gráfico. Solo lectura.|
|Leyenda.|[ChartLegend](chartlegend.md)|Representa la leyenda del gráfico. Solo lectura.|
|Series.|[ChartSeriesCollection](chartseriescollection.md)|Representa una sola serie o una colección de series del gráfico. Solo lectura.|
|title|[ChartTitle](charttitle.md)|Representa el título del gráfico especificado, incluido el texto, la visibilidad, la posición y el formato del título. Solo lectura.|

## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Elimina el objeto de gráfico.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|
|[setData(sourceData: Range or string, seriesBy: string)](#setdatasourcedata-range-or-string-seriesby-string)|void|Configura los datos de origen para el gráfico.|
|[setPosition(startCell: Range or string, endCell: Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|Coloca el gráfico con respecto a las celdas de la hoja de cálculo.|

## Detalles del método

### delete()
Elimina el objeto de gráfico.

#### Sintaxis
```js
chartObject.delete();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos
```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### Sintaxis
```js
object.load(param);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### Valores devueltos
void


### setData(sourceData: Range or string, seriesBy: string)
Configura los datos de origen para el gráfico.

#### Sintaxis
```js
chartObject.setData(sourceData, seriesBy);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|sourceData|Range or string|Dirección o nombre del intervalo que contiene los datos de origen. Si se usa una dirección o un nombre del ámbito de la hoja de cálculo, debe incluir el nombre de la hoja de cálculo (por ejemplo, "Sheet1!A5:B9"). |
|seriesBy|string|Opcional. Especifica la manera en que las columnas o las filas se usan como series de datos en el gráfico. Puede ser de una de las siguientes: Auto (valor predeterminado), Rows, Columns.  Los valores posibles son: Auto, Columns, Rows|

#### Valores devueltos
void

#### Ejemplos

Establecer `sourceData` en "A1:B4" y `seriesBy` en "Columnas"

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var sourceData = "A1:B4";
	chart.setData(sourceData, "Columns");
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### setPosition(startCell: Range or string, endCell: Range or string)
Coloca el gráfico con respecto a las celdas de la hoja de cálculo.

#### Sintaxis
```js
chartObject.setPosition(startCell, endCell);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|startCell|Range or string|Celda de inicio. Aquí es adonde se moverá el gráfico. La celda de inicio es la celda superior izquierda o superior derecha, en función de la configuración del usuario de la presentación de izquierda a derecha.|
|endCell|Range or string|Opcional. Celda final. Si se especifica, el ancho y el alto del gráfico se establecen de modo que cubran totalmente esta celda o intervalo.|

#### Valores devueltos
void

#### Ejemplos


```js
Excel.run(function (ctx) { 
	var sheetName = "Charts";
	var sourceData = sheetName + "!" + "A1:B4";
	var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", sourceData, "auto");
	chart.width = 500;
	chart.height = 300;
	chart.setPosition("C2", null);
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Ejemplos de acceso a la propiedad

Obtener un gráfico denominado "Chart1".

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.load('name');
	return ctx.sync().then(function() {
			console.log(chart.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Actualizar un gráfico, incluido el cambio de nombre, posición y tamaño.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.name="New Name";
	chart.top = 100;
	chart.left = 100;
	chart.height = 200;
	chart.weight = 200;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Asignar un nombre nuevo al gráfico y cambiar el tamaño a 200 puntos de alto y grosor. Mover Chart1 100 puntos arriba y a la izquierda. 

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
	chart.name="New Name";	
	chart.top = 100;
	chart.left = 100;
	chart.height =200;
	chart.width =200;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

