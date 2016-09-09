# Objeto Outline (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa un contenedor para objetos Paragraph.

## Properties

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|id|string|Obtiene el identificador del objeto Outline. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-id)|

## Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|pageContent|[PageContent](pagecontent.md)|Obtiene el objeto PageContent que contiene el objeto Outline. El objeto define la posición del Outline en la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-pageContent)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Obtiene la colección de objetos Paragraph en el objeto Outline. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-paragraphs)|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|Agrega el HTML especificado en la parte inferior del Outline.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|Agrega la imagen especificada en la parte inferior del Outline.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|Agrega el texto especificado en la parte inferior del Outline.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|Agrega una tabla con el número especificado de filas y columnas en la parte inferior del Outline.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendTable)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-load)|

## Detalles del método


### appendHtml(html: string)
Agrega el HTML especificado en la parte inferior del Outline.

#### Sintaxis
```js
outlineObject.appendHtml(html);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|Html|string|Cadena HTML que se anexará. Consulte el [HTML compatible](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) para la API de JavaScript de los complementos de OneNote.|

#### Valores devueltos
void

#### Ejemplos
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline")
            {
                // First item is an outline.
                outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendHtml("<p>new paragraph</p>");

                // Run the queued commands.
                return context.sync();
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### appendImage(base64EncodedImage: string, width: double, height: double)
Agrega la imagen especificada en la parte inferior del Outline.

#### Sintaxis
```js
outlineObject.appendImage(base64EncodedImage, width, height);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|base64EncodedImage|cadena|Cadena HTML que se anexará.|
|width|double|Opcional. Ancho en la unidad de puntos. El valor predeterminado es null, que respeta el ancho de la imagen.|
|height|double|Opcional. Alto en la unidad de puntos. El valor predeterminado es null, que respeta el alto de la imagen.|

#### Valores devueltos
[Image](image.md)

### appendRichText(paragraphText: string)
Agrega el texto especificado en la parte inferior del Outline.

#### Sintaxis
```js
outlineObject.appendRichText(paragraphText);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|paragraphText|string|Cadena HTML que se anexará.|

#### Valores devueltos
[RichText](richtext.md)

### appendTable(rowCount: number, columnCount: number, values: string[][])
Agrega una tabla con el número especificado de filas y columnas en la parte inferior del Outline.

#### Sintaxis
```js
outlineObject.appendTable(rowCount, columnCount, values);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|rowCount|number|Necesario. Número de filas de la tabla.|
|columnCount|number|Necesario. Número de columnas de la tabla.|
|values|string[][]|Opcional. Matriz 2D opcional. Si se especifican las cadenas correspondientes en la matriz, se rellenan las celdas.|

#### Valores devueltos
[Tabla](table.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline") {
                // First item is an outline.
                var outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendTable(2, 2, [[1, 2],[3, 4]]);

                // Run the queued commands.
                return context.sync();
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

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
