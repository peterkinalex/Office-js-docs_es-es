# Objeto FilterDatetime (API de JavaScript para Excel)

_Se aplica a: Excel 2016, Excel Online, Excel para iOS, Office 2016_

Representa cómo se filtra una fecha cuando se filtran valores.

## Propiedades

| Propiedad	   | Tipo	|Descripción
|:---------------|:--------|:----------||date|string|Fecha en formato ISO 8601 para filtrar los datos.||specificity|string|Grado de especificidad de la fecha para conservar los datos. Por ejemplo, si la fecha es 02-04-2005 y la especificidad se establece en "mes", la operación de filtrado conservará todas las filas con fecha de abril de 2005. Los valores posibles son: Year, Monday, Day, Hour, Minute, Second.|_Consulte [ejemplos](#property-access-examples) de acceso a la propiedad._

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

