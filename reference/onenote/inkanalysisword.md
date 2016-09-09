# Objeto InkAnalysisWord (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa los datos de análisis de tinta para una palabra identificada formada por trazos de tinta.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|id|string|Obtiene el identificador del objeto InkAnalysisWord. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-id)|
|languageId|string|El identificador del idioma reconocido en esta inkAnalysisWord. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-languageId)|
|wordAlternates|string|Las palabras que se han reconocido en esta palabra de tinta, por orden de probabilidad. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-wordAlternates)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|línea|[InkAnalysisLine](inkanalysisline.md)|Referencia a la InkAnalysisLine principal. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-line)|
|strokePointers|[InkStrokePointer](inkstrokepointer.md)|Referencias débiles a los trazos de tinta que se han reconocido como parte de esta palabra de análisis de tinta. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-strokePointers)|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-load)|

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
### Ejemplos de acceso a la propiedad

**wordAlternates y languageId**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    page.load('inkAnalysisOrNull/paragraphs/lines/words');
    return ctx.sync()
        .then(function() {
            var inkParagraphs = page.inkAnalysisOrNull.paragraphs;
            $.each(inkParagraphs.items, function(i, inkParagraph) {
                var inkLines = inkParagraph.lines;
                $.each(inkLines.items, function(j, inkLine) {
                    var inkWords = inkLine.words;
                    $.each(inkWords.items, function(k, inkWord) {
                    
                        // Log language Id of the word
                        console.log(inkWord.languageId);
                        
                        // Log every ink analyzed words.
                        $.each(inkWord.wordAlternates, function(l, word) {
                            console.log(word);                                  
                        })
                    })
                })
            })
        })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```