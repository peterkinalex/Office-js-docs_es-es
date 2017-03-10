# <a name="rangeview-object-javascript-api-for-excel"></a>Objeto RangeView (API de JavaScript para Excel)

RangeView representa un conjunto de celdas visibles del intervalo primario.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|cellAddresses|object[][]|Representa las direcciones de celda de RangeView. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|entero|Devuelve el número de columnas visibles. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|Representa la fórmula en notación de estilo A1.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|Representa la fórmula en notación de estilo A1, en el idioma del usuario y en la configuración regional del formato numérico. Por ejemplo, la fórmula "=SUM(A1, 1.5)" en inglés se convertiría en "=SUMME(A1; 1,5)" en alemán.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|Representa la fórmula en notación de estilo R1C1.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|index|entero|Devuelve un valor que representa el índice de RangeView. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|object[][]|Representa el código de formato numérico de Excel para la celda especificada.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|entero|Devuelve el número de filas visibles. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|text|.|Valores de texto del rango especificado. El valor Text no dependerá del ancho de la celda. La sustitución del signo # que tiene lugar en la interfaz de usuario de Excel no afectará al valor de texto devuelto por la API. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|Representa el tipo de datos de cada celda. Solo lectura. Los valores posibles son: Unknown, Empty, String, Integer, Double, Boolean, Error.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|Representa los valores sin formato de la vista del intervalo especificado. Los datos devueltos pueden ser de tipo cadena, número o booleano. La celda que contenga un error devolverá la cadena de error.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
| Relación | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|Rows|[RangeViewCollection](rangeviewcollection.md)|Representa una colección de vistas de intervalo asociadas a este. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[getRange()](#getrange)|[Range](range.md)|Obtiene el intervalo primario asociado al RangeView actual.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="getrange"></a>getRange()
Obtiene el intervalo primario asociado al RangeView actual.

#### <a name="syntax"></a>Sintaxis
```js
rangeViewObject.getRange();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Range](range.md)
