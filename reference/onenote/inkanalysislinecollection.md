# Objeto InkAnalysisLineCollection (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una colección de objetos InkAnalysisLine.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de InkAnalysisLines de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-count)|
|items|[InkAnalysisLine[]](inkanalysisline.md)|Una colección de objetos InkAnalysisLine. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-items)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkAnalysisLine](inkanalysisline.md)|Obtiene un objeto InkAnalysisLine por su identificador o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisLine](inkanalysisline.md)|Obtiene un objeto InkAnalysisLine en su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-load)|

## Detalles del método


### getItem(index: number or string)
Obtiene un objeto InkAnalysisLine por su identificador o por su índice en la colección. Solo lectura.

#### Sintaxis
```js
inkAnalysisLineCollectionObject.getItem(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number or string|El identificador del objeto InkAnalysisLine, o bien la ubicación del objeto InkAnalysisLine en la colección.|

#### Valores devueltos
[InkAnalysisLine](inkanalysisline.md)

### getItemAt(index: number)
Obtiene un objeto InkAnalysisLine en su posición en la colección.

#### Sintaxis
```js
inkAnalysisLineCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[InkAnalysisLine](inkanalysisline.md)

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
