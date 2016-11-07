# <a name="range-object-javascript-api-for-word"></a>Objeto Range (API de JavaScript para Word)

Representa un área contigua en un documento.

_Se aplica a: Word 2016, Word para iPad, Word para Mac, Word Online_

## <a name="properties"></a>Propiedades
| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|style|string|Obtiene o establece el estilo usado para el intervalo. Este es el nombre del estilo preinstalado o personalizado. En Word Online, si un nombre de estilo contiene sólo caracteres alfabéticos, el estilo *debe* ir en minúsculas, salvo pro el primer carácter, que *debe* estar en mayúsculas. Si el estilo contiene al menos un carácter no alfabético, se compara con estilos conocidos independientemente del caso y, si coinciden varios estilos, se aplicará el último estilo definido.|
|text|string|Obtiene el texto del intervalo. Solo lectura.|

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Obtiene la colección de objetos de control de contenido que se encuentran en el intervalo. Solo lectura.|
|font|[Font](font.md)|Obtiene el formato de texto del intervalo. Úselo para obtener y establecer el nombre de la fuente, el tamaño, el color y otras propiedades. Solo lectura.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Obtiene la colección de objetos inlinePicture que se encuentran en el intervalo. Solo lectura.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Obtiene la colección de objetos de párrafo que se encuentran en el intervalo. Solo lectura.|
|parentContentControl|[ContentControl](contentcontrol.md)|Obtiene el control de contenido que contiene el intervalo. Devuelve null si no hay un control de contenido principal. Solo lectura.|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Borra el contenido del objeto de intervalo. El usuario puede realizar la operación de deshacer en el contenido borrado.|
|[delete()](#delete)|void|Elimina el intervalo y su contenido del documento.|
|[getHtml()](#gethtml)|string|Obtiene la representación HTML del objeto de intervalo.|
|[getOoxml()](#getooxml)|string|Obtiene la representación OOXML del objeto de intervalo.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserta un salto en la ubicación especificada. Solo puede insertarse un salto en objetos de intervalo que se encuentran en el cuerpo principal del documento, excepto si se trata de un salto de línea, que puede insertarse en cualquier objeto de cuerpo. El valor de insertLocation puede ser 'Before' o 'After'.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Ajusta el objeto de intervalo con un control de contenido de texto enriquecido.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserta un documento en el intervalo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta HTML en el intervalo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserta una imagen en el intervalo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start', 'End', 'Before' o 'After'.
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta OOXML o wordProcessingML en el intervalo en la ubicación especificada. El valor insertLocation puede ser "Replace", "Start" o "End".|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserta un párrafo en el intervalo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserta texto en el intervalo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|Realiza una búsqueda con el valor searchOptions especificado en el ámbito del objeto de intervalo. Los resultados de la búsqueda son una colección de objetos de intervalo.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selecciona y se desplaza por la interfaz de usuario de Word hasta el intervalo. Los valores de selectionMode pueden ser 'Select', 'Start' o 'End'.|

## <a name="method-details"></a>Detalles del método

### <a name="clear"></a>clear()
Borra el contenido del objeto de intervalo. El usuario puede realizar la operación de deshacer en el contenido borrado.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.clear();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to clear the contents of the proxy range object.
    range.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the selection (range object)');
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
Elimina el intervalo y su contenido del documento.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.delete();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to delete the range object.
    range.delete();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Deleted the selection (range object)');
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
Obtiene la representación HTML del objeto de intervalo.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getHtml();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
string

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to get the HTML of the current selection.
    var html = range.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The HTML read from the document was: ' + html.value);
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
Obtiene la representación OOXML del objeto de intervalo.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.getOoxml();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
string

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to get the OOXML of the current selection.
    var ooxml = range.getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The OOXML read from the document was:  ' + ooxml.value);
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
Inserta un salto en la ubicación especificada. Solo puede insertarse un salto en objetos de intervalo que se encuentran en el cuerpo principal del documento, excepto si se trata de un salto de línea, que puede insertarse en cualquier objeto de cuerpo. El valor de insertLocation puede ser 'Before' o 'After'.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.insertBreak(breakType, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|breakType|BreakType|Necesario. Tipo de salto que se va a agregar al intervalo.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="additional-details"></a>Detalles adicionales
A excepción de los saltos de línea, no se puede insertar un salto en los objetos de encabezado, pie de página, nota al pie, nota al final, comentario y cuadro de texto.

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert a page break after the selected text.
    range.insertBreak('page', 'After');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted a page break after the selected text.');
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
Ajusta el objeto de intervalo con un control de contenido de texto enriquecido.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.insertContentControl();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[ContentControl](contentcontrol.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert a content control around the selected text,
    // and create a proxy content control object. We'll update the properties
    // on the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = "Customer-Address";
    myContentControl.title = "Enter Customer Address Here:";
    myContentControl.style = "Normal";
    myContentControl.insertText("One Microsoft Way, Redmond, WA 98052", 'replace');
    myContentControl.cannotEdit = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped a content control around the selected text.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertfilefrombase64base64file-string-insertlocation-insertlocation"></a>insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserta un documento en el intervalo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.insertFileFromBase64(base64File, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|base64File|string|Necesario. Contenido del archivo codificado en Base64 que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert base64 encoded .docx at the beginning of the range.
    // You'll need to implement getBase64() to make this work.
    range.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the range.');
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
Inserta HTML en el intervalo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|Html|string|Necesario. HTML que se va a insertar en el intervalo.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the range.');
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
Inserta una imagen en el intervalo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start', 'End', 'Before' o 'After'.

#### <a name="syntax"></a>Sintaxis
rangeObject.insertInlinePictureFromBase64(image, insertLocation);

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Necesario. Imagen codificada en base64 que se va a insertar en el intervalo.|
|insertLocation|InsertLocation|Necesario. El valor puede ser 'Replace', 'Start', 'End', 'Before' o 'After'.|

#### <a name="returns"></a>Valores devueltos
[InlinePicture](inlinepicture.md)

### <a name="insertooxmlooxml-string-insertlocation-insertlocation"></a>insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserta OOXML o wordProcessingML en el intervalo en la ubicación especificada. El valor insertLocation puede ser "Replace", "Start" o "End".

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|ooxml|string|Necesario. OOXML o wordProcessingML que se va a insertar en el intervalo.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### <a name="returns"></a>Valores devueltos
[Intervalo](range.md)

#### <a name="known-issues"></a>Problemas conocidos
Este método da como resultado una latencia elevada en Word en línea, que puede afectar a la experiencia de los usuarios de su complemento. Se recomienda utilizar este método solo cuando no hay otra solución disponible. 

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert OOXML in to the beginning of the range.
    range.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the range.');
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
Inserta un párrafo en el intervalo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.insertParagraph(paragraphText, insertLocation);
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

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert the paragraph after the range.
    range.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added to the end of the range.');
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
Inserta texto en el intervalo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start', 'End', 'Before' o 'After'.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|text|string|Necesario. Texto que se va a insertar.|
|insertLocation|InsertLocation|Obligatorio. El valor puede ser 'Replace', 'Start', 'End', 'Before' o 'After'.|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert the paragraph at the end of the range.
    range.insertText('New text inserted into the range.', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the end of the range.');
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

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to load font and style information for the range.
    context.load(range, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Show the results of the load method. Here we show the
        // property values on the range object.
        var results = "  ---Font size: " + range.font.size +
                      "  ---Font name: " + range.font.name +
                      "  ---Font color: " + range.font.color +
                      "  ---Style: " + range.style;
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

### <a name="searchsearchtext-string-searchoptions-paramtypestringssearchoptions"></a>search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Realiza una búsqueda con el valor searchOptions especificado en el ámbito del objeto de intervalo. Los resultados de la búsqueda son una colección de objetos de intervalo.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.search(searchText, searchOptions);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|searchText|string|Necesario. Texto de búsqueda.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Opcional. Opciones de la búsqueda.|

#### <a name="returns"></a>Valores devueltos
[SearchResultCollection](searchresultcollection.md)


### <a name="selectselectionmode-selectionmode"></a>select(selectionMode: SelectionMode)
Selecciona y se desplaza por la interfaz de usuario de Word hasta el intervalo. Los valores de selectionMode pueden ser 'Select', 'Start' o 'End'.

#### <a name="syntax"></a>Sintaxis
```js
rangeObject.select(selectionMode);
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

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Queue a command to select the HTML that was inserted.
    range.select();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the range.');
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
