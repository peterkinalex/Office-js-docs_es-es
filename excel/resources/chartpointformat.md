# Objeto ChartPointFormat (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Office 2016_

Representa el objeto de formato para los puntos del gráfico.

## Propiedades

Ninguno

## Relaciones
| Relación | Tipo|Descripción|
|:---------------|:--------|:----------|
|fill|[ChartFill](chartfill.md)|Representa el formato de relleno de un gráfico, que incluye información del formato de fondo. Solo lectura.|

## Métodos

| Método   | Tipo de valor devuelto|Descripción|
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
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### Valores devueltos
void

