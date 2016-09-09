# Objeto InkWord (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


El contenedor de la tinta de una palabra de un párrafo.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|id|string|Obtiene el identificador del objeto InkWord. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWord-id)|
|languageId|string|El identificador del idioma reconocido en esta palabra de tinta. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWord-languageId)|
|wordAlternates|string|Las palabras que se han reconocido en esta palabra de tinta, por orden de probabilidad. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWord-wordAlternates)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|paragraph|[Paragraph](paragraph.md)|El párrafo principal que contiene la palabra de tinta. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWord-paragraph)|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWord-load)|

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
