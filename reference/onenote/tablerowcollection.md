# Objeto TableRowCollection (API de JavaScript para Excel)

_Se aplica a: OneNote Online_  


Contiene una colección de objetos TableRow.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de filas de tabla de esta colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-count)|
|items|[TableRow[]](tablerow.md)|Colección de objetos tableRow. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-items)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[TableRow](tablerow.md)|Obtiene un objeto de fila de tabla por su id. o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|Obtiene una fila de tabla según su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-load)|

## Detalles del método


### getItem(index: number or string)
Obtiene un objeto de fila de tabla por su id. o por su índice en la colección. Solo lectura.

#### Sintaxis
```js
tableRowCollectionObject.getItem(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number o string|Número que identifica la ubicación del índice de un objeto de fila de tabla.|

#### Valores devueltos
[TableRow](tablerow.md)

### getItemAt(index: number)
Obtiene una fila de tabla según su posición en la colección.

#### Sintaxis
```js
tableRowCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[TableRow](tablerow.md)

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
