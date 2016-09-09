# Objeto InkWordCollection (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una colección de objetos InkWord.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de InkWords de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-count)|
|items|[InkWord[]](inkword.md)|Una colección de objetos InkWord. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-items)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkWord](inkword.md)|Obtiene un objeto InkWord por su identificador o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkWord](inkword.md)|Obtiene un InkWord en su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-load)|

## Detalles del método


### getItem(index: number or string)
Obtiene un objeto InkWord por su identificador o por su índice en la colección. Solo lectura.

#### Sintaxis
```js
inkWordCollectionObject.getItem(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number or string|El identificador del objeto InkWord, o bien la ubicación del índice de InkWord en la colección.|

#### Valores devueltos
[InkWord](inkword.md)

### getItemAt(index: number)
Obtiene un InkWord en su posición en la colección.

#### Sintaxis
```js
inkWordCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[InkWord](inkword.md)

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
