# Objeto ContentControl (API de JavaScript para Word)

Representa un control de contenido. Los controles de contenido son regiones delimitadas y posiblemente con etiquetas de un documento que funcionan como contenedores para tipos de contenido específicos. Los controles de contenido individuales pueden incluir contenido como imágenes, tablas o párrafos de texto con formato. Actualmente, solo se admiten controles de contenido de texto enriquecido.

_Se aplica a: Word 2016, Word para iPad, Word para Mac_

## Propiedades
| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|cannotDelete|bool|Obtiene o establece un valor que indica si el usuario puede eliminar el control de contenido. Esta propiedad y removeWhenEdited son mutuamente excluyentes.|
|cannotEdit|bool|Obtiene o establece un valor que indica si el usuario puede editar el contenido del control de contenido.|
|color|string|Obtiene o establece el color del control de contenido. El color se establece en el formato "#RRGGBB" o mediante el nombre del color.|
|placeholderText|string|Obtiene o establece el texto de marcador de posición del control de contenido. Se mostrará texto atenuado cuando el control de contenido esté vacío.|
|removeWhenEdited|bool|Obtiene o establece un valor que indica si el control de contenido se elimina después de su edición. Esta propiedad y cannotDelete son mutuamente excluyentes.|
|style|string|Obtiene o establece el estilo usado para el control de contenido. Este es el nombre del estilo preinstalado o personalizado.|
|etiqueta|string|Obtiene o establece una etiqueta para identificar un control de contenido. El complemento de ejemplo [Silly stories](https://aka.ms/sillystorywordaddin) muestra cómo se puede usar la propiedad **Tag**.|
|text|string|Obtiene el texto del control de contenido. Solo lectura.|
|title|string|Obtiene o establece el título de un control de contenido.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
| Relación | Tipo|Descripción|
|:---------------|:--------|:----------|
|appearance|**ContentControlAppearance**|Obtiene o establece el aspecto del control de contenido. El valor puede ser "boundingBox", "tags" o "hidden".|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Obtiene la colección de objetos de control de contenido que se encuentran en el control de contenido. Solo lectura.|
|font|[Font](font.md)|Obtiene el formato de texto del control de contenido. Úselo para obtener y establecer el nombre de la fuente, el tamaño, el color y otras propiedades. Solo lectura.|
|id|**[UINT]**|Obtiene un entero que representa el identificador del control de contenido. Solo lectura.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Obtiene la colección de objetos inlinePicture que se encuentran en el control de contenido. La colección no incluye imágenes flotantes. Solo lectura.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Obtiene la colección de objetos de párrafo que se encuentran en el control de contenido. Solo lectura.|
|parentContentControl|[ContentControl](contentcontrol.md)|Obtiene el control de contenido que contiene el control de contenido. Devuelve null si no hay un control de contenido principal. Solo lectura.|
|tipo|**ContentControlType**|Obtiene el tipo de control de contenido. Actualmente, solo se admiten controles de contenido de texto enriquecido. Solo lectura.|

## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Borra el contenido del control de contenido. El usuario puede realizar la operación de deshacer en el contenido borrado.|
|[delete(keepContent: bool)](#deletekeepcontent-bool)|void|Elimina el control de contenido y su contenido. Si keepContent se establece en true, el contenido no se elimina.|
|[getHtml()](#gethtml)|string|Obtiene la representación HTML del objeto de control de contenido.|
|[getOoxml()](#getooxml)|string|Obtiene la representación Office Open XML (OOXML) del objeto de control de contenido.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserta un salto en la ubicación especificada. Un salto solo puede insertarse en objetos que se encuentran en el cuerpo principal del documento, excepto si se trata de un salto de línea, que puede insertarse en cualquier objeto de cuerpo. El valor insertLocation puede ser 'Before', 'After', 'Start' o 'End'.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserta un documento en el control de contenido actual en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta HTML en el control de contenido en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserta una imagen incorporada en el control de contenido en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'. |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta OOXML o wordProcessingML en el control de contenido en la ubicación especificada. El valor insertLocation puede ser "Replace", "Start" o "End".|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Before', 'After', 'Start' o 'End'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserta texto en el control de contenido en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|Realiza una búsqueda con el valor searchOptions especificado en el ámbito del objeto de control de contenido. Los resultados de la búsqueda son una colección de objetos de intervalo.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selecciona el control de contenido. Esto hace que Word se desplace hasta la selección. El modo de selección puede ser 'Select', 'Start' o 'End'.|

## Detalles del método

### clear()
Borra el contenido del control de contenido. El usuario puede realizar la operación de deshacer en el contenido borrado.

#### Sintaxis
```js
contentControlObject.clear();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### delete(keepContent: bool)
Elimina el control de contenido y su contenido. Si keepContent se establece en true, el contenido no se elimina.

#### Sintaxis
```js
contentControlObject.delete(keepContent);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|keepContent|bool|Necesario. Indica si el contenido se debe eliminar con el control de contenido. Si keepContent se establece en true, el contenido no se elimina.|

#### Valores devueltos
void

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to delete the first content control. The
            // contents will remain in the document.
            contentControls.items[0].delete(true);
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### getHtml()
Obtiene la representación HTML del objeto de control de contenido.

#### Sintaxis
```js
contentControlObject.getHtml();
```

#### Parámetros
Ninguno

#### Valores devueltos
string

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');
    
    // Queue a command to load the tag property for all of content controls. 
    context.load(contentControlsWithTag, 'tag');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the HTML contents of the first content control.
            var html = contentControlsWithTag.items[0].getHtml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control HTML: ' + html.value);
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### getOoxml()
Obtiene la representación Office Open XML (OOXML) del objeto de control de contenido.

#### Sintaxis
```js
contentControlObject.getOoxml();
```

#### Parámetros
Ninguno

#### Valores devueltos
string

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the OOXML contents of the first content control.
            var ooxml = contentControls.items[0].getOoxml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control OOXML: ' + ooxml.value);
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserta un salto en la ubicación especificada. Un salto solo puede insertarse en objetos que se encuentran en el cuerpo principal del documento, excepto si se trata de un salto de línea, que puede insertarse en cualquier objeto de cuerpo. El valor insertLocation puede ser "Before", "After", "Start" o "End".

#### Sintaxis
```js
contentControlObject.insertBreak(breakType, insertLocation);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|breakType|BreakType|Necesario. Tipo de salto (breakType.md)|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before", "After", "Start" o "End".|

#### Valores devueltos
void

#### Detalles adicionales
A excepción de los saltos de línea, no se puede insertar un salto en objetos contenidos en encabezados, pies de página, notas al pie, notas al final, comentarios y cuadros de texto.  

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a commmand to load the id property for all of content controls. 
    context.load(contentControls, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion. We now will have 
    // access to the content control collection.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a page break after the first content control. 
            contentControls.items[0].insertBreak('page', "After");
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion. 
            return context.sync()
                .then(function () {
                    console.log('Inserted a page break after the first content control.');    
            });
        }
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
Inserta un documento en el control de contenido actual en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### Sintaxis
```js
contentControlObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|base64File|string|Necesario. Contenido codificado en Base64 del archivo que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### Valores devueltos
[Range](range.md)

### insertHtml(html: string, insertLocation: InsertLocation)
Inserta HTML en el control de contenido en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### Sintaxis
```js
contentControlObject.insertHtml(html, insertLocation);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|Html|string|Necesario. HTML que se va a insertar en el control de contenido.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### Valores devueltos
[Range](range.md)

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put HTML into the contents of the first content control.
            contentControls.items[0].insertHtml('<strong>HTML content inserted into the content control.</strong>', 'Start');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted HTML in the first content control.');
            });
        }
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
Inserta una imagen incorporada en el control de contenido en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### Sintaxis
contentControlObject.insertInlinePictureFromBase64(image, insertLocation);

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Necesario. Imagen codificada en base64 que se va a insertar en el control de contenido.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### Valores devueltos
[InlinePicture](inlinepicture.md)



### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserta OOXML o wordProcessingML en el control de contenido en la ubicación especificada. El valor insertLocation puede ser "Replace", "Start" o "End".

#### Sintaxis
```js
contentControlObject.insertOoxml(ooxml, insertLocation);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|ooxml|string|Necesario. OOXML o wordProcessingML que se va a insertar en el control de contenido.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### Valores devueltos
[Range](range.md)

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put OOXML into the contents of the first content control.
            contentControls.items[0].insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", "End");
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted OOXML in the first content control.');
            });
        }
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
Lea [Crear complementos mejores para Word con Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx) para obtener instrucciones sobre cómo trabajar con OOXML.

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Before', 'After', 'Start' o 'End'.

#### Sintaxis
```js
contentControlObject.insertParagraph(paragraphText, insertLocation);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|paragraphText|string|Necesario. Texto de párrafo que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before", "After", "Start" o "End".|

#### Valores devueltos
[Paragraph](paragraph.md)

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a paragraph after the first content control. 
            contentControls.items[0].insertParagraph('Text of the inserted paragraph.', 'After');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted a paragraph after the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertText(text: string, insertLocation: InsertLocation)
Inserta texto en el control de contenido en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.

#### Sintaxis
```js
contentControlObject.insertText(text, insertLocation);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|texto|string|Necesario. Texto que se va a insertar en el control de contenido.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Replace", "Start" o "End".|

#### Valores devueltos
[Range](range.md)

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to replace text in the first content control. 
            contentControls.items[0].insertText('Replaced text in the first content control.', 'Replace');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Replaced text in the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

El complemento de ejemplo [Silly stories](https://aka.ms/sillystorywordaddin) muestra cómo se usa el método **insertText**.

### load(param: object)
Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.

#### Sintaxis
```js
object.load(param);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### Valores devueltos
void

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy range object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to create the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = 'Customer-Address';
    myContentControl.title = ' has t';
    myContentControl.style = 'Heading 2';
    myContentControl.insertText('One Microsoft Way, Redmond, WA 98052', 'replace');
    myContentControl.cannotEdit = true;
    myContentControl.appearance = 'tags';
    
    // Queue a command to load the id property for the content control you created.
    context.load(myContentControl, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Created content control with id: ' + myContentControl.id);
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
Realiza una búsqueda con el valor searchOptions especificado en el ámbito del objeto de control de contenido. Los resultados de la búsqueda son una colección de objetos de intervalo.

#### Sintaxis
```js
contentControlObject.search(searchText, searchOptions);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|searchText|string|Necesario. Texto de búsqueda.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Opcional. Opciones de la búsqueda.|

#### Valores devueltos
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
Selecciona el control de contenido. Esto hace que Word se desplace hasta la selección. El modo de selección puede ser 'Select', 'Start' o 'End'.

#### Sintaxis
```js
contentControlObject.select(selectionMode);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Opcional. El modo de selección puede ser 'Select', 'Start' o 'End'. 'Select' es el valor predeterminado.|

#### Valores devueltos
void

#### Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to select the first content control.
            contentControls.items[0].select();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Selected the first content control.');
            });
        }
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

### Cargar todas las propiedades del control de contenido.
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to load the properties on the first content control. 
            contentControls.items[0].load(  'appearance,' +
                                            'cannotDelete,' +
                                            'cannotEdit,' +
                                            'color,' +
                                            'id,' +
                                            'placeHolderText,' +
                                            'removeWhenEdited,' +
                                            'title,' +
                                            'text,' +
                                            'type,' +
                                            'style,' +
                                            'tag,' +
                                            'font/size,' +
                                            'font/name,' +
                                            'font/color');             
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Property values of the first content control:' + 
                        '   ----- appearance: ' + contentControls.items[0].appearance + 
                        '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
                        '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
                        '   ----- color: ' + contentControls.items[0].color +
                        '   ----- id: ' + contentControls.items[0].id +
                        '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
                        '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
                        '   ----- title: ' + contentControls.items[0].title +
                        '   ----- text: ' + contentControls.items[0].text +
                        '   ----- type: ' + contentControls.items[0].type +
                        '   ----- style: ' + contentControls.items[0].style +
                        '   ----- tag: ' + contentControls.items[0].tag +
                        '   ----- font size: ' + contentControls.items[0].font.size +
                        '   ----- font name: ' + contentControls.items[0].font.name +
                        '   ----- font color: ' + contentControls.items[0].font.color);
            });
        }
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

Use el [conjunto de requisitos](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) en las comprobaciones en tiempo de ejecución para asegurarse de que la aplicación es compatible con la versión de host de Word. Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 
