# <a name="paragraph-object-javascript-api-for-word"></a>Objeto Paragraph (API de JavaScript para Word)

Representa un solo párrafo de una selección, intervalo, control de contenido o cuerpo del documento.

_Se aplica a: Word 2016, Word para iPad, Word para Mac, Word Online_

## <a name="properties"></a>Propiedades
| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|outlineLevel|int|Obtiene o establece el nivel de esquema del párrafo.|
|style|string|Obtiene o establece el estilo usado en el párrafo. Este es el nombre del estilo preinstalado o personalizado. En el ejemplo [Word-Add-in-DocumentAssembly][paragraph.style], se muestra cómo se puede establecer el estilo de párrafo.|
|text|string|Obtiene el texto del párrafo. Solo lectura.|

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|alignment|**Alignment**|Obtiene o establece la alineación de un párrafo. El valor puede ser "left", "centered", "right" o "justified".|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Obtiene la colección de objetos de control de contenido que se encuentran en el párrafo. Solo lectura.|
|firstLineIndent|**float**|Obtiene o establece el valor (en puntos) para una sangría en la primera línea o francesa. Para establecer una sangría en la primera línea se debe usar un valor positivo, para establecer una sangría francesa se debe usar un valor negativo.|
|font|[Font](font.md)|Obtiene el formato de texto del párrafo. Úselo para obtener y establecer el nombre de la fuente, el tamaño, el color y otras propiedades. Solo lectura.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Obtiene la colección de objetos inlinePicture que se encuentran en el párrafo. La colección no incluye imágenes flotantes. Solo lectura.|
|leftIndent|**float**|Obtiene o establece el valor de sangría izquierda (en puntos) del párrafo.|
|lineSpacing|**float**|Obtiene o establece el espaciado de línea (en puntos) del párrafo especificado. En la interfaz de usuario de Word, este valor se divide entre 12.|
|lineUnitAfter|**float**|Obtiene o establece la cantidad de espaciado (en líneas de cuadrícula) después del párrafo.|
|lineUnitBefore|**float**|Obtiene o establece la cantidad de espaciado (en líneas de cuadrícula) antes del párrafo.|
|parentContentControl|[ContentControl](contentcontrol.md)|Obtiene el control de contenido que contiene el párrafo. Devuelve null si no hay un control de contenido principal. Solo lectura.|
|rightIndent|**float**|Obtiene o establece el valor de sangría derecha (en puntos) del párrafo.|
|spaceAfter|**float**|Obtiene o establece el espaciado (en puntos) después del párrafo.|
|spaceBefore|**float**|Obtiene o establece el espaciado (en puntos) antes del párrafo.|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Borra el contenido del objeto de párrafo. El usuario puede realizar la operación de deshacer en el contenido borrado.|
|[delete()](#delete)|void|Elimina el párrafo y su contenido del documento.|
|[getHtml()](#gethtml)|string|Obtiene la representación HTML del objeto de párrafo.|
|[getOoxml()](#getooxml)|string|Obtiene la representación Office Open XML (OOXML) del objeto de párrafo.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserta un salto en la ubicación especificada. Un salto solo puede insertarse en párrafos que se encuentran en el cuerpo principal del documento, excepto si se trata de un salto de línea, que puede insertarse en cualquier objeto de cuerpo. El valor insertLocation puede ser "After" o "Before".|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Ajusta el objeto de párrafo con un control de contenido de texto enriquecido.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserta un documento en el párrafo actual en la ubicación especificada. El valor insertLocation puede ser 'Start' o 'End'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta HTML en el párrafo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserta una imagen en el párrafo en la ubicación especificada. El valor insertLocation puede ser 'Before', 'After', 'Start' o 'End'.|
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta OOXML o wordProcessingML en el párrafo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserta texto en el párrafo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|Realiza una búsqueda con el valor searchOptions especificado en el ámbito del objeto de párrafo. Los resultados de la búsqueda son una colección de objetos de intervalo.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selecciona y se desplaza por la interfaz de usuario de Word hasta el párrafo. El modo de selección puede ser 'Select', 'Start' o 'End'. 'Select' es el valor predeterminado.|

## <a name="method-details"></a>Detalles del método

### <a name="clear"></a>clear()
Borra el contenido del objeto de párrafo. El usuario puede realizar la operación de deshacer en el contenido borrado.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.clear();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to clear the contents of the first paragraph.
        paragraphs.items[0].clear();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Cleared the contents of the first paragraph.');
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

### <a name="delete"></a>delete()
Elimina el párrafo y su contenido del documento.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.delete();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the text property for all of the paragraphs.
    context.load(paragraphs, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to delete the first paragraph.
        paragraphs.items[0].delete();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Deleted the first paragraph.');
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

### <a name="gethtml"></a>getHtml()
Obtiene la representación HTML del objeto de párrafo.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.getHtml();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
string

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a a set of commands to get the HTML of the first paragraph.
        var html = paragraphs.items[0].getHtml();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
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

### <a name="getooxml"></a>getOoxml()
Obtiene la representación Office Open XML (OOXML) del objeto de párrafo.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.getOoxml();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
string

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a a set of commands to get the OOXML of the first paragraph.
        var ooxml = paragraphs.items[0].getOoxml();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph OOXML: ' + ooxml.value);
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

### <a name="insertbreakbreaktype-breaktype-insertlocation-insertlocation"></a>insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserta un salto en la ubicación especificada. Un salto solo puede insertarse en párrafos que se encuentran en el cuerpo principal del documento, excepto si se trata de un salto de línea, que puede insertarse en cualquier objeto de cuerpo. El valor insertLocation puede ser "Before" o "After".

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.insertBreak(breakType, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|breakType|BreakType|Necesario. Tipo de salto que se va a agregar al documento.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="additional-details"></a>Detalles adicionales
No puede insertar un salto en encabezados, pies de página, notas al pie, notas al final, comentarios y cuadros de texto.

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert a page break after the first paragraph.
        paragraph.insertBreak('page', 'After');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a page break after the paragraph.');
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

### <a name="insertcontentcontrol"></a>insertContentControl()
Ajusta el objeto de párrafo con un control de contenido de texto enriquecido.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.insertContentControl();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[ContentControl](contentcontrol.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to wrap the first paragraph in a rich text content control.
        paragraph.insertContentControl();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Wrapped the first paragraph in a content control.');
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

#### <a name="additional-information"></a>Información adicional
En el ejemplo [Word-Add-in-DocumentAssembly][paragraph.insertContentControl], se muestra cómo se puede establecer el método insertContentControl.

### <a name="insertfilefrombase64base64file-string-insertlocation-insertlocation"></a>insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserta un documento en el párrafo actual en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.insertFileFromBase64(base64File, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|base64File|string|Necesario. Contenido del archivo codificado en Base64 que se va a insertar.|
|insertLocation|InsertLocation|Obligatorio. El valor puede ser 'Replace', 'Start' o 'End'.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert base64 encoded .docx at the beginning of the first paragraph.
        // This won't work unless you have a definition for getBase64().
        paragraph.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted base64 encoded content at the beginning of the first paragraph.');
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

### <a name="inserthtmlhtml-string-insertlocation-insertlocation"></a>insertHtml(html: string, insertLocation: InsertLocation)
Inserta HTML en el párrafo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|Html|string|Necesario. HTML que se va a insertar en el párrafo.|
|insertLocation|InsertLocation|Obligatorio. El valor puede ser 'Replace', 'Start' o 'End'.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert HTML content at the end of the first paragraph.
        paragraph.insertHtml('<strong>Inserted HTML.</strong>', Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted HTML content at the end of the first paragraph.');
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

### <a name="insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation"></a>insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserta una imagen en el párrafo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Necesario. HTML que se va a insertar en el párrafo.|
|insertLocation|InsertLocation|Obligatorio. El valor puede ser 'Replace', 'Start' o 'End'.|

#### <a name="returns"></a>Valores devueltos
[InlinePicture](inlinepicture.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        var b64encodedImg = "iVBORw0KGgoAAAANSUhEUgAAAB4AAAANCAIAAAAxEEnAAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACFSURBVDhPtY1BEoQwDMP6/0+XgIMTBAeYoTqso9Rkx1zG+tNj1H94jgGzeNSjteO5vtQQuG2seO0av8LzGbe3anzRoJ4ybm/VeKEerAEbAUpW4aWQCmrGFWykRzGBCnYy2ha3oAIq2MloW9yCCqhgJ6NtcQsqoIKdjLbFLaiACnYyf2fODbrjZcXfr2F4AAAAAElFTkSuQmCC";

        // Queue a command to insert a base64 encoded image at the beginning of the first paragraph.
        paragraph.insertInlinePictureFromBase64(b64encodedImg, Word.InsertLocation.start);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added an image to the first paragraph.');
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

#### <a name="additional-information"></a>Información adicional
En el ejemplo [Word-Add-in-DocumentAssembly][paragraph.insertpicture], se proporciona otro ejemplo de cómo se puede insertar una imagen en un párrafo.

### <a name="insertooxmlooxml-string-insertlocation-insertlocation"></a>insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserta OOXML o wordProcessingML en el párrafo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|ooxml|string|Necesario. OOXML o wordProcessingML que se va a insertar en el párrafo.|
|insertLocation|InsertLocation|Obligatorio. El valor puede ser 'Replace', 'Start' o 'End'.|

#### <a name="returns"></a>Valores devueltos
[Intervalo](range.md)

#### <a name="known-issues"></a>Problemas conocidos
Este método da como resultado una latencia elevada en Word en línea, que puede afectar a la experiencia de los usuarios de su complemento. Se recomienda utilizar este método solo cuando no hay otra solución disponible. 

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert Ooxml content into the first paragraph.
        var ooxmlContent = "<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
        paragraph.insertOoxml(ooxmlContent, Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted OOXML at the end of the first paragraph.');
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

#### <a name="additional-information"></a>Información adicional
Lea [Crear complementos mejores para Word con Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx) para obtener instrucciones sobre cómo trabajar con OOXML.

### <a name="insertparagraphparagraphtext-string-insertlocation-insertlocation"></a>insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|paragraphText|string|Necesario. Texto de párrafo que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### <a name="returns"></a>Valores devueltos
[Paragraph](paragraph.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert the paragraph after the current paragraph.
        paragraph.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a new paragraph at the end of the first paragraph.');
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

### <a name="inserttexttext-string-insertlocation-insertlocation"></a>insertText(text: string, insertLocation: InsertLocation)
Inserta texto en el párrafo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|text|string|Necesario. Texto que se va a insertar.|
|insertLocation|InsertLocation|Obligatorio. El valor puede ser 'Replace', 'Start' o 'End'.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert text into the end of the paragraph.
        paragraph.insertText('New text inserted into the paragraph.', Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted text at the end of the first paragraph.');
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

### <a name="loadparam-object"></a>load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### <a name="syntax"></a>Sintaxis
```js
object.load(param);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to load font information for the paragraph.
        context.load(paragraph, 'font/size, font/name, font/color');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            // Show the results of the load method. Here we show the
            // property values on the paragraph object. Note that we
            // requested the style property in the first load command.
            var results = "<strong>Paragraph</strong>--" +
                          "--Font size: " + paragraph.font.size +
                          "--Font name: " + paragraph.font.name +
                          "--Font color: " + paragraph.font.color +
                          "--Style: " + paragraph.style;

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

### <a name="searchsearchtext-string-searchoptions-paramtypestringssearchoptions"></a>search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Realiza una búsqueda con el valor searchOptions especificado en el ámbito del objeto de párrafo. Los resultados de la búsqueda son una colección de objetos de intervalo.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.search(searchText, searchOptions);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|searchText|string|Necesario. Texto de búsqueda.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Opcional. Opciones de la búsqueda.|

#### <a name="returns"></a>Valores devueltos
[SearchResultCollection](searchresultcollection.md)

### <a name="selectselectionmode-selectionmode"></a>select(selectionMode: SelectionMode)
Selecciona y se desplaza por la interfaz de usuario de Word hasta el párrafo.

#### <a name="syntax"></a>Sintaxis
```js
paragraphObject.select(selectionMode);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Opcional. El modo de selección puede ser 'Select', 'Start' o 'End'. 'Select' es el valor predeterminado.|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the last paragraph a create a
        // proxy paragraph object.
        var paragraph = paragraphs.items[paragraphs.items.length - 1];

        // Queue a command to select the paragraph. The Word UI will
        // move to the selected paragraph.
        paragraph.select();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Selected the last paragraph.');
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

## <a name="support-details"></a>Detalles de compatibilidad
Use el [conjunto de requisitos](../office-add-in-requirement-sets.md) en las comprobaciones en tiempo de ejecución para asegurarse de que la aplicación es compatible con la versión de host de Word. Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


[paragraph.insertContentControl]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L161 "insertar control de contenido"
[paragraph.style]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L172 "establecer estilo"
[paragraph.insertpicture]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L236 "insertar imagen"
