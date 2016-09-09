# Objeto NotebookCollection (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una colección de blocs de notas.

## Properties

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de blocs de notas de una colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-count)|
|items|[Notebook[]](notebook.md)|Una colección de objetos de bloc de notas. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-items)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[NotebookCollection](notebookcollection.md)|Obtiene la colección de blocs de notas con el nombre especificado que están abiertos en la instancia de la aplicación.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getByName)|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Bloc de notas](notebook.md)|Obtiene un bloc de notas por ID o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Bloc de notas](notebook.md)|Obtiene un bloc de notas según su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-load)|

## Detalles del método


### getByName(name: string)
Obtiene la colección de blocs de notas con el nombre especificado que están abiertos en la instancia de la aplicación.

#### Sintaxis
```js
notebookCollectionObject.getByName(name);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|name|string|El nombre del bloc de notas.|

#### Valores devueltos
[NotebookCollection](notebookcollection.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // Get the notebooks that are open in the application instance and have the specified name.
    var notebooks = context.application.notebooks.getByName("Homework");

    // Queue a command to load the notebooks. 
    // For best performance, request specific properties.           
    notebooks.load("id,name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index, for example: notebooks.items[0]
            if (notebooks.items.length > 0) {
                console.log("Notebook name: " + notebooks.items[0].name);
                console.log("Notebook ID: " + notebooks.items[0].id);
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

### getItem(index: number or string)
Obtiene un bloc de notas por id. o por su índice en la colección. Solo lectura.

#### Sintaxis
```js
notebookCollectionObject.getItem(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number o string|El id. del bloc de notas, o bien la ubicación del índice del bloc de notas en la colección.|

#### Valores devueltos
[Bloc de notas](notebook.md)

### getItemAt(index: number)
Obtiene un bloc de notas según su posición en la colección.

#### Sintaxis
```js
notebookCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[Bloc de notas](notebook.md)

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
### Ejemplos de acceso a la propiedad

**Items**
```js
OneNote.run(function (context) {

    // Get the notebooks that are open in the application instance and have the specified name.
    var notebooks = context.application.notebooks.getByName("Homework");

    // Queue a command to load the notebooks. 
    // For best performance, request specific properties.           
    notebooks.load("id");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index, for example: notebooks.items[0]
            $.each(notebooks.items, function(index, notebook) {
                notebook.addSection("Biology");
                notebook.addSection("Spanish");
                notebook.addSection("Computer Science");
            });
            
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

