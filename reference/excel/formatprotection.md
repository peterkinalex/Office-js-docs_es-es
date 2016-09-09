# Objeto FormatProtection (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS y Office 2016_

Representa la protección de formato de un objeto range.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|formulaHidden|bool|Indica si Excel oculta la fórmula de las celdas del rango. Un valor nulo indica que el intervalo no tiene una configuración de ajuste uniforme.|
|locked|bool|Indica si Excel bloquea las celdas del objeto. Un valor nulo indica que todo el rango no tiene una configuración de bloqueo uniforme.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


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
