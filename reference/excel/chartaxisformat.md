# Objeto ChartAxisFormat (API de JavaScript para Excel)

Encapsula las propiedades de formato del eje del gráfico.

## Propiedades

Ninguno

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|font|[ChartFont](chartfont.md)|Representa los atributos de fuente (por ejemplo, nombre de fuente, tamaño de fuente, color, etc.) de un elemento del eje del gráfico. Solo lectura.|
|line|[ChartLineFormat](chartlineformat.md)|Representa el formato de línea de gráfico. Solo lectura.|

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
