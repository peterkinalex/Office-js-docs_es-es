# Objeto SectionCollection (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una colección de secciones.

## Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de secciones de una colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-count)|
|items|[Section[]](section.md)|Colección de objetos de sección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-items)|

_Consulte los [ejemplos](#ejemplos) de acceso a la propiedad._

## Relaciones
Ninguno


## Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[SectionCollection](sectioncollection.md)|Obtiene la colección de secciones con el nombre especificado.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getByName)|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Sección](section.md)|Obtiene una sección por su id. o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Sección](section.md)|Obtiene una sección según su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-load)|

## Detalles del método


### getByName(name: string)
Obtiene la colección de secciones con el nombre especificado.

#### Sintaxis
```js
sectionCollectionObject.getByName(name);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|name|string|El nombre de la sección.|

#### Valores devueltos
[SectionCollection](sectioncollection.md)

#### Ejemplos
```js
OneNote.run(function (context) {

    // Get the sections in the current notebook.
    var sections = context.application.getActiveNotebook().sections;

    // Queue a command to load the sections. 
    // For best performance, request specific properties.
    sections.load("id"); 
    
    // Get the sections with the specified name.
    var groceriesSections = sections.getByName("Groceries");
    
    // Queue a command to load the sections with the specified name.
    groceriesSections.load("id,name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index.
            if (groceriesSections.items.length > 0) {
                console.log("Section name: " + groceriesSections.items[0].name);
                console.log("Section ID: " + groceriesSections.items[0].id);
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
Obtiene una sección por id. o por su índice en la colección. Solo lectura.

#### Sintaxis
```js
sectionCollectionObject.getItem(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number o string|El id. de la sección, o bien la ubicación del índice de la sección en la colección.|

#### Valores devueltos
[Sección](section.md)

### getItemAt(index: number)
Obtiene una sección según su posición en la colección.

#### Sintaxis
```js
sectionCollectionObject.getItemAt(index);
```

#### Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### Valores devueltos
[Sección](section.md)

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

    // Get the sections in the current notebook.
    var sections = context.application.getActiveNotebook().sections;

    // Queue a command to load the sections. 
    // For best performance, request specific properties.
    sections.load("name"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Iterate through the collection or access items individually by index, for example: sections.items[0]
            $.each(sections.items, function(index, section) {
                if (section.name === "Homework") {
                    section.addPage("Biology");
                    section.addPage("Spanish");
                    section.addPage("Computer Science");
                }
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

