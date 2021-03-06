# <a name="section-object-(javascript-api-for-word)"></a>Objeto Section (API de JavaScript para Word)

Representa una sección de un documento de Word.

_Se aplica a: Word 2016, Word para iPad, Word para Mac, Word Online_

## <a name="properties"></a>Propiedades
Ninguno

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción|
|:---------------|:--------|:----------|
|body|[Body](body.md)|Obtiene el cuerpo de la sección. Esto no incluye el encabezado y pie de página y otros metadatos de la sección. Solo lectura.|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[getFooter(type: HeaderFooterType)](#getfootertype-headerfootertype)|[Body](body.md)|Obtiene uno de los pies de página de la sección.|
|[getHeader(type: HeaderFooterType)](#getheadertype-headerfootertype)|[Body](body.md)|Obtiene uno de los encabezados de la sección.|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## <a name="method-details"></a>Detalles del método

### <a name="getfooter(type:-headerfootertype)"></a>getFooter(type: HeaderFooterType)
Obtiene uno de los pies de página de la sección.

#### <a name="syntax"></a>Sintaxis
```js
sectionObject.getFooter(type);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|tipo|HeaderFooterType|Necesario. Tipo de pie de página que se va a devolver. Este valor puede ser "primary", "firstPage" o "evenPages".|

#### <a name="returns"></a>Valores devueltos
[Body](body.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;
    
    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object the primary footer of the first section. 
        // Note that the footer is a body object.
        var myFooter = mySections.items[0].getFooter("primary");
        
        // Queue a command to insert text at the end of the footer.
        myFooter.insertText("This is a footer.", Word.InsertLocation.end);
        
        // Queue a command to wrap the header in a content control.
        myFooter.insertContentControl();
                              
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a footer to the first section.");
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
### <a name="getheader(type:-headerfootertype)"></a>getHeader(type: HeaderFooterType)
Obtiene uno de los encabezados de la sección.

#### <a name="syntax"></a>Sintaxis
```js
sectionObject.getHeader(type);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|tipo|HeaderFooterType|Necesario. Tipo de encabezado que se va a devolver. Este valor puede ser "primary", "firstPage" o "evenPages".|

#### <a name="returns"></a>Valores devueltos
[Body](body.md)

#### <a name="examples"></a>Ejemplos
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;
    
    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object the primary header of the first section. 
        // Note that the header is a body object.
        var myHeader = mySections.items[0].getHeader("primary");
        
        // Queue a command to insert text at the end of the header.
        myHeader.insertText("This is a header.", Word.InsertLocation.end);
        
        // Queue a command to wrap the header in a content control.
        myHeader.insertContentControl();
                              
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a header to the first section.");
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
