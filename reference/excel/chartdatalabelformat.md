# <a name="chartdatalabelformat-object-(javascript-api-for-excel)"></a>Objeto ChartDataLabelFormat (API de JavaScript para Excel)

Encapsula las propiedades de formato de las etiquetas de datos del gráfico.

## <a name="properties"></a>Propiedades

Ninguno

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|fill|[ChartFill](chartfill.md)|Representa el formato de relleno de la etiqueta de datos del gráfico actual. Solo lectura.|
|font|[ChartFont](chartfont.md)|Representa los atributos de fuente (nombre de fuente, tamaño de fuente, color, etc.) de una etiqueta de datos del gráfico. Solo lectura.|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## <a name="method-details"></a>Detalles del método


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