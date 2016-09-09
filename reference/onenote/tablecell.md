# Objeto TableCell (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una celda en una tabla de OneNote.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|cellIndex|int|Obtiene el índice de la celda de la fila. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-cellIndex)|
|id|string|Obtiene el identificador de la celda. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-id)|
|rowIndex|int|Obtiene el índice de fila de la celda en la tabla. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-rowIndex)|
|shadingColor|string|Obtiene y establece el color de sombreado de la celda.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-shadingColor)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Obtiene la colección de objetos Paragraph en el objeto TableCell. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-paragraphs)|
|parentRow|[TableRow](tablerow.md)|Obtiene la fila primaria de la celda. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-parentRow)|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|Agrega el HTML especificado en la parte inferior de TableCell.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|Agrega la imagen especificada a la celda de la tabla.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|Agrega el texto especificado a la celda de la tabla.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|Agrega una tabla con el número especificado de filas y columnas a la celda de la tabla.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendTable)|
|[clear()](#clear)|void|Borra el contenido de la celda.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-clear)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-load)|

## Detalles del método


### appendHtml(html: string)
Agrega el HTML especificado en la parte inferior de TableCell.

#### Sintaxis
```js
tableCellObject.appendHtml(html);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|Html|string|Cadena HTML que se anexará. Consulte el [HTML compatible](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) para la API de JavaScript de los complementos de OneNote.|

#### Valores devueltos
void

#### Ejemplos
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two and add "Hello".
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                cell.appendHtml("<p>Hello</p>");
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});


### appendImage(base64EncodedImage: string, width: double, height: double)
Adds the specified image to table cell.

#### Syntax
```js
tableCellObject.appendImage(base64EncodedImage, width, height);
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
Agrega el texto especificado a la celda de la tabla.

#### Sintaxis
```js
tableCellObject.appendRichText(paragraphText);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|paragraphText|string|Cadena HTML que se anexará.|

#### Valores devueltos
[RichText](richtext.md)

#### Ejemplos
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    var appendedRichText = null;
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two and add "Hello".
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                appendedRichText = cell.appendRichText("Hello");
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### appendTable(rowCount: number, columnCount: number, values: string[][])
Agrega una tabla con el número especificado de filas y columnas a la celda de la tabla.

#### Sintaxis
```js
tableCellObject.appendTable(rowCount, columnCount, values);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|rowCount|number|Necesario. Número de filas de la tabla.|
|columnCount|number|Necesario. Número de columnas de la tabla.|
|values|string[][]|Opcional. Matriz 2D opcional. Si se especifican las cadenas correspondientes en la matriz, se rellenan las celdas.|

#### Valores devueltos
[Table](table.md)

### clear()
Borra el contenido de la celda.

#### Sintaxis
```js
tableCellObject.clear();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

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
**id, cellIndex, rowIndex**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                
                // Queue a command to load the table cell.
                ctx.load(cell);
                ctx.sync().then(function() {
                    console.log("Cell Id: " + cell.id);
                    console.log("Cell Index: " + cell.cellIndex);
                    console.log("Cell's Row Index: " + cell.rowIndex);
                });
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**parentTable, cells**
```js
ParentTable, ParentRow, Paragraphs
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                
                // Queue a command to load parentTable, parentRow and paragraphs of the table cell.
                ctx.load(cell, "parentTable, parentRow, paragraphs");
                
                ctx.sync().then(function() {
                    console.log("Parent Table Id: " + cell.parentTable.id);
                    console.log("Parent Row Id: " + cell.parentRow.id);
                    var paragraphs = cell.paragraphs;
                    
                    for (var i = 0; i < paragraphs.items.length; i++) {
                        console.log("Paragraph Id: " + paragraphs.items[i].id);
                    }
                });
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

