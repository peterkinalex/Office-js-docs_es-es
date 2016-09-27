# Objeto SearchOptions (API de JavaScript para Word)

Especifica las opciones que se van a incluir en una operación de búsqueda.

_Se aplica a: Word 2016, Word para iPad, Word para Mac_

## Properties
| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|ignorePunct|bool|Obtiene o establece un valor que indica si se van a pasar por alto todos los caracteres de puntuación entre las palabras. Corresponde a la casilla Omitir puntuación en el cuadro de diálogo Buscar y reemplazar.|
|ignoreSpace|bool|Obtiene o establece un valor que indica si se van a pasar por alto todos los espacios en blanco entre las palabras. Corresponde a la casilla Omitir espacios en blanco en el cuadro de diálogo Buscar y reemplazar.|
|matchCase|bool|Obtiene o establece un valor que indica si se va a realizar una búsqueda distinguiendo entre mayúsculas y minúsculas. Corresponde a la casilla Coincidir mayúsculas y minúsculas en el cuadro de diálogo Buscar y reemplazar (menú Edición).|
|matchPrefix|bool|Obtiene o establece un valor que indica si se van a buscar palabras que empiecen por la cadena de búsqueda. Corresponde a la casilla Coincidir prefijo en el cuadro de diálogo Buscar y reemplazar.|
|matchSoundsLike|bool|**Esta opción quedó en desuso en la actualización de junio de 2016**. Obtiene o establece un valor que indica si se van a buscar palabras que se parezcan a la cadena de búsqueda. Corresponde a la casilla Se parece a en el cuadro de diálogo Buscar y reemplazar.|
|matchSuffix|bool|Obtiene o establece un valor que indica si se van a buscar palabras que finalicen por la cadena de búsqueda. Corresponde a la casilla Coincidir sufijo en el cuadro de diálogo Buscar y reemplazar.|
|matchWholeWord|bool|Obtiene o establece un valor que indica si se van a buscar solamente palabras completas y no texto que forme parte de una palabra más larga. Corresponde a la casilla Solo palabras completas en el cuadro de diálogo Buscar y reemplazar.|
|matchWildCards|bool|Obtiene o establece un valor que indica si la búsqueda se realizará usando operadores de búsqueda especiales. Corresponde a la casilla Usar caracteres comodín en el cuadro de diálogo Buscar y reemplazar.|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

Las opciones de búsqueda son opcionales. Las opciones de búsqueda deben especificarse en todos los métodos de búsqueda mediante un literal de objeto:

```js
    search('searchstring', {searchOption1:bool, ...searchOptionN:bool}
```

Puede proporcionar una o más propiedades de opción de búsqueda en el literal de objeto para especificar las opciones de búsqueda. 

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|

## Detalles del método

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

## Ejemplos de acceso a la propiedad

### Búsqueda Omitir puntuación
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document and ignore punctuation.
    var searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Búsqueda basada en un prefijo
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document based on a prefix.
    var searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Búsqueda basada en un sufijo
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document for any string of characters after 'ly'.
    var searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'orange';
            searchResults.items[i].font.highlightColor = 'black';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Búsqueda con caracteres comodín
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    var searchResults = context.document.body.search('to*n', {matchWildCards: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = 'pink';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


## Guía de caracteres comodín 

| Para buscar:         | Carácter comodín |  Ejemplo |
|:-----------------|:--------|:----------|
| Cualquier carácter| ? |s?t busca sat y set. |
|Cualquier cadena de caracteres| * |s*d busca sad y started.|
|Principio de palabra|< |<(inter) busca interesting e intercept, pero no splintered.|
|Final de palabra |> |(in)> busca in y within, pero no interesting.|
|Uno de los caracteres especificados|[ ] |w[io]n busca win y won.|
|Cualquier carácter individual en este intervalo| [-] |[r-t]ight busca right y sight. Los intervalos deben estar en orden ascendente.|
|Cualquier carácter individual excepto los caracteres que aparecen en el intervalo dentro de los corchetes|[!x-z] |t[!a-m]ck busca tock y tuck, pero no tack ni tick.|
|Exactamente n apariciones del carácter o expresión anteriores|{n} |fe\{2\}d busca feed pero no fed.|
|Como mínimo n apariciones del carácter o expresión anteriores|{n,} |fe{1,}d busca fed y feed.|
|De n a m apariciones del carácter o expresión anteriores|{n,m} |10{1,3} busca 10, 100 y 1000.|
|Una o más apariciones del carácter o expresión anteriores|@ |lo@t busca lot y loot.|


## Detalles de compatibilidad
Use el [conjunto de requisitos](../office-add-in-requirement-sets.md) en las comprobaciones en tiempo de ejecución para asegurarse de que la aplicación es compatible con la versión de host de Word. Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).
