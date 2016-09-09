# Objeto InkAnalysisParagraphCollection (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una colección de objetos InkAnalysisParagraph.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de InkAnalysisParagraphs de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-count)|
|items|[InkAnalysisParagraph[]](inkanalysisparagraph.md)|Una colección de objetos InkAnalysisParagraph. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-items)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkAnalysisParagraph](inkanalysisparagraph.md)|Obtiene un objeto InkAnalysisParagraph por su identificador o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisParagraph](inkanalysisparagraph.md)|Obtiene un objeto InkAnalysisParagraph en su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-load)|

## Detalles del método


### getItem(index: number or string)
Obtiene un objeto InkAnalysisParagraph por su identificador o por su índice en la colección. Solo lectura.

#### Sintaxis
```js
inkAnalysisParagraphCollectionObject.getItem(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number or string|El identificador del objeto InkAnalysisParagraph, o bien la ubicación del objeto InkAnalysisParagraph en la colección.|

#### Valores devueltos
[InkAnalysisParagraph](inkanalysisparagraph.md)

### getItemAt(index: number)
Obtiene un objeto InkAnalysisParagraph en su posición en la colección.

#### Sintaxis
```js
inkAnalysisParagraphCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[InkAnalysisParagraph](inkanalysisparagraph.md)

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
