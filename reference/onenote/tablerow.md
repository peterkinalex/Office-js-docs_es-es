# Objeto TableRow (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una fila de una tabla.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|cellCount|int|Obtiene el número de celdas del intervalo. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-cellCount)|
|id|string|Obtiene el identificador de la fila. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-id)|
|rowIndex|int|Obtiene el índice de la fila en la tabla primaria. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-rowIndex)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|cells|[TableCellCollection](tablecellcollection.md)|Obtiene las celdas de la fila. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-cells)|
|parentTable|[Table](table.md)|Obtiene la tabla primaria. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-parentTable)|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[clear()](#clear)|void|Borra el contenido de la fila.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-clear)|
|[insertRowAsSibling(insertLocation: string, values: string[])](#insertrowassiblinginsertlocation-string-values-string)|[TableRow](tablerow.md)|Inserta una fila antes o después de la fila actual.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-insertRowAsSibling)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-load)|
|[setShadingColor(colorCode: string)](#setshadingcolorcolorcode-string)|void|Establece el color de sombreado de todas las celdas de la fila.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-setShadingColor)|

## Detalles del método


### clear()
Borra el contenido de la fila.

#### Sintaxis
```js
tableRowObject.clear();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

### insertRowAsSibling(insertLocation: string, values: string[])
Inserta una fila antes o después de la fila actual.

#### Sintaxis
```js
tableRowObject.insertRowAsSibling(insertLocation, values);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|insertLocation|string|Dónde se deben insertar las filas nuevas con respecto a la fila actual.  Los valores posibles son: Before, After|
|values|string[]|Opcional. Cadenas para insertar en la nueva fila, especificadas como una matriz. No debe tener más celdas que la fila actual. Opcional.|

#### Valores devueltos
[TableRow](tablerow.md)

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
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load table.rows.
                ctx.load(table, "rows");
                
                // Run the queued commands
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    rows.items[1].insertRowAsSibling("Before", ["cell0", "cell1"]);
                    return ctx.sync();
                });
            }
        }
    })
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

### setShadingColor(colorCode: string)
Establece el color de sombreado de todas las celdas de la fila.

#### Sintaxis
```js
tableRowObject.setShadingColor(colorCode);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|colorCode|string|El código de color que hay que establecer en las celdas./param|

#### Valores devueltos
void
### Ejemplos de acceso a la propiedad
**id, cellCount, rowIndex**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load table.rows.
                ctx.load(table, "rows");
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    
                    // for each table row, log cell count and row index.
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Id: " + rows.items[i].id);
                        console.log("Row " + i + " Cell Count: " + rows.items[i].cellCount);
                        console.log("Row " + i + " Row Index: " + rows.items[i].rowIndex);
                    }
                    return ctx.sync();
                });
            }
        }
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
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load parentTable and cells of each row in the table.
                ctx.load(table, "rows/parentTable, rows/cells");
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    
                    // for each row, log parentTable and cells
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Parent Table Id: " + rows.items[i].parentTable.id);
                        var cells = rows.items[i].cells;
                        for (var j = 0 ; j < cells.items.length; j++) {
                            console.log("Row " + i + " Cell " + j + " Id: " + cells.items[j].id);
                        }
                    }
                    return ctx.sync();
                });
            }
        }
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

