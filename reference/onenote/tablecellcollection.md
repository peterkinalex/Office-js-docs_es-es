# Objeto TableCellCollection (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Contiene una colección de objetos TableCell.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de tableCell de esta colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-count)|
|items|[TableCell[]](tablecell.md)|Colección de objetos tableCell. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-items)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[TableCell](tablecell.md)|Obtiene un objeto TableCell por su id. o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[TableCell](tablecell.md)|Obtiene una celda basada en su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-load)|

## Detalles del método


### getItem(index: number or string)
Obtiene un objeto TableCell por su id. o por su índice en la colección. Solo lectura.

#### Sintaxis
```js
tableCellCollectionObject.getItem(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number o string|Número que identifica la ubicación del índice de un objeto de celda de tabla.|

#### Valores devueltos
[TableCell](tablecell.md)

### getItemAt(index: number)
Obtiene una celda basada en su posición en la colección.

#### Sintaxis
```js
tableCellCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[TableCell](tablecell.md)

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
