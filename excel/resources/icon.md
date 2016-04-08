# Objeto Icon (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS, Office 2016_

Representa un icono de celda.

## Propiedades

| Propiedad	   | Tipo	|Descripción
|:---------------|:--------|:----------||index|int|Representa el índice del icono del conjunto dado.||set|string|Representa el conjunto al que pertenece el icono. Los valores posibles son: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|_Consulte [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
Ninguna


## Métodos

| Método		   | Tipo de valor devuelto	|Descripción||:---------------|:--------|:----------||[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método


## load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

### Sintaxis
```js
object.load(param);
```

### Parámetros
| Parámetro	   | Tipo	|Descripción||:---------------|:--------|:----------||param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

### Valores devueltos
void

