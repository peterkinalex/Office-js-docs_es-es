# Objeto InlinePicture (API de JavaScript para Word)

Representa una imagen incorporada.

_Se aplica a: Word 2016, Word para iPad, Word para Mac_

## Propiedades
| Propiedad   | Tipo|Descripción
|:---------------|:--------|:----------|
|altTextDescription|string|Obtiene o establece una cadena que representa el texto alternativo asociado a la imagen incorporada.|
|altTextTitle|string|Obtiene o establece una cadena que contiene el título de la imagen incorporada.|
|hyperlink|string|Obtiene o establece el hipervínculo asociado a la imagen incorporada.|
|lockAspectRatio|bool|Obtiene o establece un valor que indica si la imagen incorporada mantiene sus proporciones originales cuando se cambia su tamaño.|


_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## Relaciones
| Relación | Tipo|Descripción|
|:---------------|:--------|:----------|
|height|**float**|Obtiene o establece un número que describe la altura de la imagen incorporada. Se mide en puntos. |
|parentContentControl|[ContentControl](contentcontrol.md)|Obtiene el control de contenido que contiene la imagen incorporada. Devuelve null si no hay un control de contenido principal. Solo lectura.|
|Paragraph|[paragraph](paragraph.md)|Obtiene el párrafo que contiene la imagen incorporada. Solo lectura.
|width|**float**|Obtiene o establece un número que describe el ancho de la imagen incorporada. Se mide en puntos.|

## Métodos

| Método   | Tipo de valor devuelto|Descripción|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Elimina la imagen del documento.|
|[getBase64ImageSrc()](#getbase64imagesrc)|string|Obtiene un objeto cuyo valor es la representación de cadena codificada en Base64 de la imagen incorporada.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserta un salto en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Ajusta la imagen incorporada con un control de contenido de texto enriquecido.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserta un documento en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta HTML en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserta una imagen en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Before' o 'After'. |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta OOXML en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserta texto en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selecciona la imagen y se desplaza por la interfaz de usuario de Word hasta ella. Los valores de selectionMode pueden ser 'Select', 'Start' o 'End'.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método

### delete()
Elimina la imagen del documento.

#### Sintaxis
```js
inlinePictureObject.delete();
```

#### Parámetros
Ninguno

#### Valores devueltos
void

### getBase64ImageSrc()
Obtiene un objeto cuyo valor es la representación de cadena codificada en Base64 de la imagen incorporada.

#### Sintaxis
```js
inlinePictureObject.getBase64ImageSrc();
```

#### Parámetros
Ninguno

#### Valores devueltos
string

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)

#### Sintaxis
```js
inlinePictureObject.insertBreak(breakType, insertLocation);
```
#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|breakType|BreakType|Necesario. Tipo de salto que se va a agregar al cuerpo.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### Valores devueltos
void

### insertContentControl()
Ajusta la imagen incorporada con un control de contenido de texto enriquecido.

#### Sintaxis
```js
inlinePictureObject.insertContentControl();
```

#### Parámetros
Ninguno

#### Valores devueltos
[ContentControl](contentcontrol.md)

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserta un documento en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### Sintaxis
```js
inlinePictureObject.insertFileFromBase64(base64File, insertLocation);
```
#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|base64File|string|Necesario. Contenido codificado en base64 de un archivo docx.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### Valores devueltos
[Range](range.md)

### insertHtml(html: string, insertLocation: InsertLocation)
Inserta HTML en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### Sintaxis
```js
inlinePictureObject.insertHtml(html, insertLocation);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|Html|string|Necesario. HTML que se va a insertar en el documento.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### Valores devueltos
[Range](range.md)


### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserta una imagen en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### Sintaxis
inlinePictureObject.insertInlinePictureFromBase64(image, insertLocation);

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Necesario. Imagen codificada en base64 que se va a insertar en el cuerpo.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### Valores devueltos
[InlinePicture](inlinepicture.md)


### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserta OOXML en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### Sintaxis
```js
inlinePictureObject.insertOoxml(ooxml, insertLocation);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|ooxml|string|Necesario. OOXML que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### Valores devueltos
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### Sintaxis
```js
inlinePictureObject.insertParagraph(paragraphText, insertLocation);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|paragraphText|string|Necesario. Texto de párrafo que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### Valores devueltos
[Paragraph](paragraph.md)

### insertText(text: string, insertLocation: InsertLocation)
Inserta texto en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### Sintaxis
```js
inlinePictureObject.insertText(text, insertLocation);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|texto|string|Necesario. Texto que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### Valores devueltos
[Range](range.md)

### select(selectionMode: SelectionMode)
Selecciona la imagen y se desplaza por la interfaz de usuario de Word hasta ella. Los valores de selectionMode pueden ser 'Select', 'Start' o 'End'.

#### Sintaxis
```js
inlinePictureObject.select(selectionMode);
```

#### Parámetros
| Parámetro   | Tipo|Descripción|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Opcional. El modo de selección puede ser 'Select', 'Start' o 'End'. 'Select' es el valor predeterminado.|

#### Valores devueltos
void

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

## Detalles de compatibilidad

Use el [conjunto de requisitos](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) en las comprobaciones en tiempo de ejecución para asegurarse de que la aplicación es compatible con la versión de host de Word. Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 
