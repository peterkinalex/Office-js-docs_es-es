# Objeto ChartAxes (API de JavaScript para Excel)

Representa los ejes del gráfico.

## Propiedades

Ninguno

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|categoryAxis|[ChartAxis](chartaxis.md)|Representa el eje de categorías de un gráfico. Solo lectura.|
|seriesAxis|[ChartAxis](chartaxis.md)|Representa el eje de series de un gráfico tridimensional. Solo lectura.|
|valueAxis|[ChartAxis](chartaxis.md)|Representa el eje de valores de un eje. Solo lectura.|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método


### load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### Sintaxis
```js
object.load(param);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### Valores devueltos
void
