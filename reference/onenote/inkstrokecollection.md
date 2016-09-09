# Objeto InkStrokeCollection (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_   


Representa una colección de objetos InkStroke.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de InkStrokes de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-count)|
|items|[InkStroke[]](inkstroke.md)|Una colección de objetos InkStroke. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-items)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkStroke](inkstroke.md)|Obtiene un objeto InkStroke por su identificador o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkStroke](inkstroke.md)|Obtiene un InkStroke según su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-load)|

## Detalles del método


### getItem(index: number or string)
Obtiene un objeto InkStroke por su identificador o por su índice en la colección. Solo lectura.

#### Sintaxis
```js
inkStrokeCollectionObject.getItem(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number or string|El identificador del objeto InkStroke, o bien la ubicación del índice de InkStroke en la colección.|

#### Valores devueltos
[InkStroke](inkstroke.md)

### getItemAt(index: number)
Obtiene un InkStroke según su posición en la colección.

#### Sintaxis
```js
inkStrokeCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[InkStroke](inkstroke.md)

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
