# Objeto InkStrokePointer (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Referencia débil a un objeto de trazo de tinta y su elemento principal de contenido

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|contentId|string|Representa el identificador del objeto de contenido de página correspondiente a este trazo.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-contentId)|
|inkStrokeId|string|Representa el identificador del trazo de tinta.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-inkStrokeId)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-load)|

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
