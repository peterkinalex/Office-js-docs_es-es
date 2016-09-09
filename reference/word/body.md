# Objeto Body (API de JavaScript para Word)

Representa el cuerpo de un documento o una sección.

_Se aplica a: Word 2016, Word para iPad, Word para Mac_

## Properties
| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|style|string|Obtiene o establece el estilo usado para el cuerpo. Este es el nombre del estilo preinstalado o personalizado.|
|text|string|Obtiene el texto del cuerpo. Use el método insertText para insertar texto. Solo lectura.|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Obtiene la colección de objetos de control de contenido de texto enriquecido que se encuentran en el cuerpo. Solo lectura.|
|font|[Fuente](font.md)|Obtiene el formato de texto del cuerpo. Úselo para obtener y establecer el nombre de la fuente, el tamaño, el color y otras propiedades. Solo lectura.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Obtiene la colección de objetos inlinePicture que se encuentran en el cuerpo. La colección no incluye imágenes flotantes. Solo lectura.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Obtiene la colección de objetos de párrafo que se encuentran en el cuerpo. Solo lectura.|
|parentContentControl|[ContentControl](contentcontrol.md)|Obtiene el control de contenido que contiene el cuerpo. Devuelve null si no hay un control de contenido principal. Solo lectura.|

## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Borra el contenido del objeto de cuerpo. El usuario puede realizar la operación de deshacer en el contenido borrado.|
|[getHtml()](#gethtml)|string|Obtiene la representación HTML del objeto de cuerpo.|
|[getOoxml()](#getooxml)|cadena|Obtiene la representación OOXML (Office Open XML) del objeto de cuerpo.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserta un salto en la ubicación especificada. Un salto solo puede insertarse en el cuerpo principal del documento, excepto si se trata de un salto de línea, que puede insertarse en cualquier objeto de cuerpo. El valor insertLocation puede ser "Start" o "End".|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Ajusta el objeto de cuerpo con un control de contenido de texto enriquecido.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserta un documento en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta HTML en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserta una imagen en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Start' o 'End'. |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta OOXML en la ubicación especificada.  El valor insertLocation puede ser “Replace”, “Start” o “End”.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Start' o 'End'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserta texto en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Realiza una búsqueda con el valor searchOptions especificado en el ámbito del objeto de cuerpo. Los resultados de la búsqueda son una colección de objetos de intervalo.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selecciona el cuerpo y se desplaza por la interfaz de usuario de Word hasta él. Los valores de selectionMode pueden ser 'Select', 'Start' o 'End'.|

## Detalles del método

### clear()
Borra el contenido del objeto de cuerpo. El usuario puede realizar la operación de deshacer en el contenido borrado.

#### Sintaxis
```js
bodyObject.clear();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to clear the contents of the body.
    body.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the body contents.');
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

El ejemplo de complemento [Silly stories](https://aka.ms/sillystorywordaddin) muestra cómo se usa el método **clear** para borrar el contenido de un documento.

### getHtml()
Obtiene la representación HTML del objeto de cuerpo.

#### Sintaxis
```js
bodyObject.getHtml();
```

#### Parámetros
Ninguno

#### Valores devueltos
string

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the HTML contents of the body.
    var bodyHTML = body.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body HTML contents: " + bodyHTML.value);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### getOoxml()
Obtiene la representación OOXML (Office Open XML) del objeto de cuerpo.

#### Sintaxis
```js
bodyObject.getOoxml();
```

#### Parámetros
Ninguno

#### Valores devueltos
string

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the OOXML contents of the body.
    var bodyOOXML = body.getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body OOXML contents: " + bodyOOXML.value);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserta un salto en la ubicación especificada. Un salto solo puede insertarse en el cuerpo principal del documento, excepto si se trata de un salto de línea, que puede insertarse en cualquier objeto de cuerpo. El valor insertLocation puede ser "Start" o "End".

#### Sintaxis
```js
bodyObject.insertBreak(breakType, insertLocation);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|breakType|BreakType|Necesario. Tipo de salto que se va a agregar al cuerpo.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Start" o "End".|

#### Valores devueltos
void

#### Detalles adicionales
A excepción de los saltos de línea, no se puede insertar un salto en encabezados, pies de página, notas al pie, notas al final, comentarios y cuadros de texto.

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (ctx) {

    // Create a proxy object for the document body.
    var body = ctx.document.body;

    // Queue a commmand to insert a page break at the start of the document body.
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        console.log('Added a page break at the start of the document body.');
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### insertContentControl()
Ajusta el objeto de cuerpo con un control de contenido de texto enriquecido.

#### Sintaxis
```js
bodyObject.insertContentControl();
```

#### Parámetros
Ninguno

#### Valores devueltos
[ContentControl](contentcontrol.md)

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to wrap the body in a content control.
    body.insertContentControl();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped the body in a content control.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserta un documento en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### Sintaxis
```js
bodyObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|base64File|string|Necesario. Contenido del archivo codificado en Base64 que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### Valores devueltos
[Range](range.md)

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert base64 encoded .docx at the beginning of the content body.
    // You will need to implement getBase64() to pass in a string of a base64 encoded docx file.
    body.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

El ejemplo de complemento [Silly stories](https://aka.ms/sillystorywordaddin) muestra cómo se puede usar el método **insertFileFromBase64** para insertar archivos docx desde un servicio.

### insertHtml(html: string, insertLocation: InsertLocation)
Inserta HTML en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### Sintaxis
```js
bodyObject.insertHtml(html, insertLocation);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|Html|string|Necesario. HTML que se va a insertar en el documento.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### Valores devueltos
[Range](range.md)

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert HTML in to the beginning of the body.
    body.insertHtml('<strong>This is text inserted with body.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserta una imagen en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Start' o 'End'.

#### Sintaxis
bodyObject.insertInlinePictureFromBase64(image, insertLocation);

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Necesario. Imagen codificada en base64 que se va a insertar en el cuerpo.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Start" o "End".|

#### Valores devueltos
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserta OOXML en la ubicación especificada.  El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### Sintaxis
```js
bodyObject.insertOoxml(ooxml, insertLocation);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|ooxml|string|Necesario. OOXML o wordProcessingML que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### Valores devueltos
[Range](range.md)

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert OOXML in to the beginning of the body.
    body.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### Información adicional
Lea [Crear mejores complementos para Word con Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx) para obtener instrucciones sobre cómo trabajar con OOXML. El ejemplo [Word-Add-in-DocumentAssembly][body.insertOoxml] muestra cómo puede usar esta API para ensamblar un documento.

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Start' o 'End'.

#### Sintaxis
```js
bodyObject.insertParagraph(paragraphText, insertLocation);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|paragraphText|string|Necesario. Texto de párrafo que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Start" o "End".|

#### Valores devueltos
[Paragraph](paragraph.md)

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    body.insertParagraph('Content of a new paragraph', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added at the end of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### Información adicional
El ejemplo [Word-Add-in-DocumentAssembly][body.insertParagraph] muestra cómo puede usar el método insertParagraph para ensamblar un documento.

### insertText(text: string, insertLocation: InsertLocation)
Inserta texto en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### Sintaxis
```js
bodyObject.insertText(text, insertLocation);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|text|string|Necesario. Texto que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### Valores devueltos
[Range](range.md)

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert text in to the beginning of the body.
    body.insertText('This is text inserted with body.insertText()', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
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

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Realiza una búsqueda con las opciones de búsqueda especificadas en el ámbito del objeto de cuerpo. Los resultados de la búsqueda son una colección de objetos de intervalo.

#### Sintaxis
```js
bodyObject.search(searchText, searchOptions);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|searchText|string|Necesario. Texto de búsqueda.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Opcional. Opciones de la búsqueda.|

#### Valores devueltos
[SearchResultCollection](searchresultcollection.md)

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to search the document.
    var searchResults = context.document.body.search('video', {matchCase: false});

    // Queue a commmand to load the results.
    context.load(searchResults, 'text, font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        var results = 'Found count: ' + searchResults.items.length +
                      '; we highlighted the results.';

        // Queue a command to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].font.color = '#FF0000'    // Change color to Red
          searchResults.items[i].font.highlightColor = '#FFFF00';
          searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log(results);
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### Información adicional
El ejemplo [Word-Add-in-DocumentAssembly][body.search] proporciona otro ejemplo sobre cómo buscar en un documento.

### select(selectionMode: SelectionMode)
Selecciona el cuerpo y se desplaza por la interfaz de usuario de Word hasta él. Los valores de selectionMode pueden ser 'Select', 'Start' o 'End'.

#### Sintaxis
```js
bodyObject.select(selectionMode);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Opcional. El modo de selección puede ser 'Select', 'Start' o 'End'. 'Select' es el valor predeterminado.|

#### Valores devueltos
void

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to select the document body. The Word UI will
    // move to the selected document body.
    body.select();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## Ejemplos de acceso a la propiedad

### Obtener la propiedad Text del objeto de cuerpo.
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load the text in document body.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### Obtener el estilo y el tamaño de fuente, el nombre de la fuente y las propiedades de color de fuente del objeto de cuerpo.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## Detalles de compatibilidad

Use el [conjunto de requisitos](../office-add-in-requirement-sets.md) en las comprobaciones en tiempo de ejecución para asegurarse de que la aplicación es compatible con la versión de host de Word. Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


[body.insertOoxml]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L127 "insert OOXML"
[body.insertParagraph]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L153 "insert paragraph"
[body.search]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L261 "body search"
