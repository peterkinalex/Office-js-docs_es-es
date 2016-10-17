# <a name="inlinepicture-object-(javascript-api-for-word)"></a>Objeto InlinePicture (API de JavaScript para Word)

Representa una imagen incorporada.

_Se aplica a: Word 2016, Word para iPad, Word para Mac, Word Online_

## <a name="properties"></a>Propiedades
| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|altTextDescription|string|Obtiene o establece una cadena que representa el texto alternativo asociado a la imagen incorporada.|
|altTextTitle|string|Obtiene o establece una cadena que contiene el título de la imagen incorporada.|
|hyperlink|string|Obtiene o establece el hipervínculo asociado a la imagen incorporada.|
|lockAspectRatio|bool|Obtiene o establece un valor que indica si la imagen incorporada mantiene sus proporciones originales cuando se cambia su tamaño.|

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|height|**float**|Obtiene o establece un número que describe la altura de la imagen incorporada. Este valor se mide en puntos. |
|parentContentControl|[ContentControl](contentcontrol.md)|Obtiene el control de contenido que contiene la imagen incorporada. Devuelve null si no hay un control de contenido principal. Solo lectura.|
|paragraph|[paragraph](paragraph.md)|Obtiene el párrafo que contiene la imagen incorporada. Solo lectura.
|width|**float**|Obtiene o establece un número que describe el ancho de la imagen incorporada. Este valor se mide en puntos.|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Elimina la imagen del documento.|
|[getBase64ImageSrc()](#getbase64imagesrc)|object|Obtiene un objeto cuyo valor es la representación de cadena codificada en Base64 de la imagen incorporada.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserta un salto en la ubicación especificada. El valor de insertLocation puede ser 'Before' o 'After'.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Ajusta la imagen incorporada con un control de contenido de texto enriquecido.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserta un documento en el cuerpo en la ubicación especificada. El valor de insertLocation puede ser 'Before' o 'After'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta HTML en la ubicación especificada. El valor de insertLocation puede ser 'Before' o 'After'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserta una imagen en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Before' o 'After'. |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserta OOXML en la ubicación especificada.  El valor de insertLocation puede ser 'Before' o 'After'.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserta texto en el cuerpo en la ubicación especificada. El valor de insertLocation puede ser 'Before' o 'After'.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selecciona la imagen y se desplaza por la interfaz de usuario de Word hasta ella. Los valores de selectionMode pueden ser 'Select', 'Start' o 'End'.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## <a name="method-details"></a>Detalles del método

### <a name="delete()"></a>delete()
Elimina la imagen del documento.

#### <a name="syntax"></a>Sintaxis
```js
inlinePictureObject.delete();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void

### <a name="getbase64imagesrc()"></a>getBase64ImageSrc()
Obtiene un objeto cuyo valor es la representación de cadena codificada en Base64 de la imagen incorporada.

#### <a name="syntax"></a>Sintaxis
```js
var base64 = inlinePictureObject.getBase64ImageSrc();
return context.sync().then(function () {    
    console.log("base64 string is " + base64.value);
});

```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
object 



### <a name="insertbreak(breaktype:-breaktype,-insertlocation:-insertlocation)"></a>insertBreak(breakType: BreakType, insertLocation: InsertLocation)

#### <a name="syntax"></a>Sintaxis
```js
inlinePictureObject.insertBreak(breakType, insertLocation);
```
#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|breakType|BreakType|Necesario. Tipo de salto que se va a agregar al cuerpo.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### <a name="returns"></a>Valores devueltos
void

### <a name="insertcontentcontrol()"></a>insertContentControl()
Ajusta la imagen incorporada con un control de contenido de texto enriquecido.

#### <a name="syntax"></a>Sintaxis
```js
inlinePictureObject.insertContentControl();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[ContentControl](contentcontrol.md)

### <a name="insertfilefrombase64(base64file:-string,-insertlocation:-insertlocation)"></a>insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserta un documento en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### <a name="syntax"></a>Sintaxis
```js
inlinePictureObject.insertFileFromBase64(base64File, insertLocation);
```
#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|base64File|string|Necesario. Contenido codificado en base64 de un archivo docx.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

### <a name="inserthtml(html:-string,-insertlocation:-insertlocation)"></a>insertHtml(html: string, insertLocation: InsertLocation)
Inserta HTML en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### <a name="syntax"></a>Sintaxis
```js
inlinePictureObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|Html|string|Necesario. HTML que se va a insertar en el documento.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)


### <a name="insertinlinepicturefrombase64(base64encodedimage:-string,-insertlocation:-insertlocation)"></a>insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserta una imagen en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### <a name="syntax"></a>Sintaxis
inlinePictureObject.insertInlinePictureFromBase64(image, insertLocation);

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Necesario. Imagen codificada en base64 que se va a insertar en el cuerpo.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### <a name="returns"></a>Valores devueltos
[InlinePicture](inlinepicture.md)


### <a name="insertooxml(ooxml:-string,-insertlocation:-insertlocation)"></a>insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserta OOXML en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### <a name="syntax"></a>Sintaxis
```js
inlinePictureObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|ooxml|string|Necesario. OOXML que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

### <a name="insertparagraph(paragraphtext:-string,-insertlocation:-insertlocation)"></a>insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### <a name="syntax"></a>Sintaxis
```js
inlinePictureObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|paragraphText|string|Necesario. Texto de párrafo que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### <a name="returns"></a>Valores devueltos
[Paragraph](paragraph.md)

### <a name="inserttext(text:-string,-insertlocation:-insertlocation)"></a>insertText(text: string, insertLocation: InsertLocation)
Inserta texto en el cuerpo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.

#### <a name="syntax"></a>Sintaxis
```js
inlinePictureObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|text|string|Necesario. Texto que se va a insertar.|
|insertLocation|InsertLocation|Necesario. El valor puede ser "Before" o "After".|

#### <a name="returns"></a>Valores devueltos
[Range](range.md)

### <a name="select(selectionmode:-selectionmode)"></a>select(selectionMode: SelectionMode)
Selecciona la imagen y se desplaza por la interfaz de usuario de Word hasta ella. Los valores de selectionMode pueden ser 'Select', 'Start' o 'End'.

#### <a name="syntax"></a>Sintaxis
```js
inlinePictureObject.select(selectionMode);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Opcional. El modo de selección puede ser 'Select', 'Start' o 'End'. 'Select' es el valor predeterminado.|

#### <a name="returns"></a>Valores devueltos
void

### <a name="load(param:-object)"></a>load(param: object)
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

## <a name="support-details"></a>Detalles de compatibilidad
Use el [conjunto de requisitos](../office-add-in-requirement-sets.md) en las comprobaciones en tiempo de ejecución para asegurarse de que la aplicación es compatible con la versión de host de Word. Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).
