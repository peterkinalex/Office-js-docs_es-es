# Objeto Table (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una tabla en una página de OneNote.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|borderVisible|bool|Obtiene o establece si los bordes están visibles o no. True si son visibles, false si están ocultos.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-borderVisible)|
|columnCount|int|Obtiene el número de columnas de la tabla. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-columnCount)|
|id|string|Obtiene el identificador de la tabla. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-id)|
|rowCount|int|Obtiene el número de filas de la tabla. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-rowCount)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|paragraph|[Paragraph](paragraph.md)|Obtiene el objeto Paragraph que contiene el objeto Table. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-paragraph)|
|rows|[TableRowCollection](tablerowcollection.md)|Obtiene todas las filas de la tabla. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-rows)|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[appendColumn(values: string[])](#appendcolumnvalues-string)|void|Agrega una columna al final de la tabla. Los valores, si se especifican, se establecen en la nueva columna. De lo contrario, la columna está vacía.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-appendColumn)|
|[appendRow(values: string[])](#appendrowvalues-string)|[TableRow](tablerow.md)|Agrega una fila al final de la tabla. Los valores, si se especifican, se establecen en la nueva fila. De lo contrario, la fila está vacía.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-appendRow)|
|[clear()](#clear)|void|Borra el contenido de la tabla.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-clear)|
|[getCell(rowIndex: number, cellIndex: number)](#getcellrowindex-number-cellindex-number)|[TableCell](tablecell.md)|Obtiene la celda de tabla de una fila y columna especificadas.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-getCell)|
|[insertColumn(index: number, values: string[])](#insertcolumnindex-number-values-string)|void|Inserta una columna en el índice especificado de la tabla. Los valores, si se especifican, se establecen en la nueva columna. De lo contrario, la columna está vacía.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-insertColumn)|
|[insertRow(index: number, values: string[])](#insertrowindex-number-values-string)|[TableRow](tablerow.md)|Inserta una fila en el índice especificado de la tabla. Los valores, si se especifican, se establecen en la nueva fila. De lo contrario, la fila está vacía.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-insertRow)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-load)|
|[setShadingColor(colorCode: string)](#setshadingcolorcolorcode-string)|void|Establece el color de sombreado de todas las celdas de la tabla.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-setShadingColor)|

## Detalles del método


### appendColumn(values: string[])
Agrega una columna al final de la tabla. Los valores, si se especifican, se establecen en la nueva columna. De lo contrario, la columna está vacía.

#### Sintaxis
```js
tableObject.appendColumn(values);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|values|string[]|Opcional. Opcional. Cadenas para insertar en la nueva columna, especificadas como una matriz. No debe tener más valores que filas de la tabla.|

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
        
        // for each table, append a column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                table.appendColumn(["cell0", "cell1"]);
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


### appendRow(values: string[])
Agrega una fila al final de la tabla. Los valores, si se especifican, se establecen en la nueva fila. De lo contrario, la fila está vacía.

#### Sintaxis
```js
tableObject.appendRow(values);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|values|string[]|Opcional. Opcional. Cadenas para insertar en la nueva fila, especificadas como una matriz. No debe tener más valores que columnas de la tabla.|

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
        
        // for each table, append a column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var row = table.appendRow(["cell0", "cell1"]);
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


### clear()
Borra el contenido de la tabla.

#### Sintaxis
```js
tableObject.clear();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

### getCell(rowIndex: number, cellIndex: number)
Obtiene la celda de tabla de una fila y columna especificadas.

#### Sintaxis
```js
tableObject.getCell(rowIndex, cellIndex);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|rowIndex|number|Índice de la fila.|
|cellIndex|number|Índice de la celda de la fila.|

#### Valores devueltos
[TableCell](tablecell.md)

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
        
        // for each table, get a cell in the second row and third column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(2 /*Row Index*/, 3 /*Column Index*/);
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


### insertColumn(index: number, values: string[])
Inserta una columna en el índice especificado de la tabla. Los valores, si se especifican, se establecen en la nueva columna. De lo contrario, la columna está vacía.

#### Sintaxis
```js
tableObject.insertColumn(index, values);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Índice donde se insertará la columna en la tabla.|
|values|string[]|Opcional. Opcional. Cadenas para insertar en la nueva columna, especificadas como una matriz. No debe tener más valores que filas de la tabla.|

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
        
        // for each table, insert a column at index two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                table.insertColumn(2, ["cell0", "cell1"]);
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


### insertRow(index: number, values: string[])
Inserta una fila en el índice especificado de la tabla. Los valores, si se especifican, se establecen en la nueva fila. De lo contrario, la fila está vacía.

#### Sintaxis
```js
tableObject.insertRow(index, values);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Índice donde se insertará la fila en la tabla.|
|values|string[]|Opcional. Opcional. Cadenas para insertar en la nueva fila, especificadas como una matriz. No debe tener más valores que columnas de la tabla.|

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
        
        // for each table, insert a row at index two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var row = table.insertRow(2, ["cell0", "cell1"]);
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
Establece el color de sombreado de todas las celdas de la tabla.

#### Sintaxis
```js
tableObject.setShadingColor(colorCode);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|colorCode|string|El código de color que hay que establecer en las celdas./param|

#### Valores devueltos
void
### Ejemplos de acceso a la propiedad
**columnCount, rowCount, id**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // For each table, log properties.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                ctx.load(table);
                return ctx.sync().then(function() {
                    console.log("Table Id: " + table.id);
                    console.log("Row Count: " + table.rowCount);
                    console.log("Column Count: " + table.columnCount);
                    return ctx.sync();
                });
            }
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

**paragraph, rows**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, log its paragraph id.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                ctx.load(table, "paragraph/id, rows/id");
                return ctx.sync().then(function() {
                    console.log("Paragraph Id: " + table.paragraph.id);
                    var rows = table.rows;
                    
                    // for each rows in the table, log row index and id.
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Id: " + rows.items[i].id);
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

