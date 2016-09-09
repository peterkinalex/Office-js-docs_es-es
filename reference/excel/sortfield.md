# Objeto SortField (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS y Office 2016_

Representa una condición en una operación de ordenación.

## Propiedades

| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|ascending|bool|Representa si la ordenación se realiza en orden ascendente.|
|color|string|Representa el color que es el destino de la condición si la ordenación se realiza según la fuente o el color de celda.|
|dataOption|string|Representa opciones de ordenación adicionales para este campo. Los valores posibles son: Normal, TextAsNumber.|
|key|int|Representa la columna (o fila, según la orientación de ordenación) en que se encuentra la condición. Se representa como un desplazamiento de la primera columna (o fila).|
|sortOn|string|Representa el tipo de ordenación de esta condición. Los valores posibles son: Value, CellColor, FontColor, Icon.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|icono|[Icono](icon.md)|Representa el icono que es el destino de la condición si la ordenación se realiza según el icono de la celda.|

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
