# <a name="sectiongroupcollection-object-(javascript-api-for-onenote)"></a>Objeto SectionGroupCollection (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una colección de grupos de secciones.

## <a name="properties"></a>Properties

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de grupos de secciones de una colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-count)|
|items|[SectionGroup[]](sectiongroup.md)|Una colección de objetos sectionGroup. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-items)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[SectionGroupCollection](sectiongroupcollection.md)|Obtiene la colección de grupos de secciones con el nombre especificado.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getByName)|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[SectionGroup](sectiongroup.md)|Obtiene un grupo de secciones por su identificador o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[SectionGroup](sectiongroup.md)|Obtiene un grupo de secciones según su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-load)|

## <a name="method-details"></a>Detalles del método


### <a name="getbyname(name:-string)"></a>getByName(name: string)
Obtiene la colección de grupos de secciones con el nombre especificado.

#### <a name="syntax"></a>Sintaxis
```js
sectionGroupCollectionObject.getByName(name);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|name|string|El nombre del grupo de secciones.|

#### <a name="returns"></a>Valores devueltos
[SectionGroupCollection](sectiongroupcollection.md)

#### <a name="examples"></a>Ejemplos
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups. 
    // For best performance, request specific properties.
    sectionGroups.load("id"); 

    // Get the section groups with the specified name.
    var labsSectionGroups = sectionGroups.getByName("Labs");

    // Queue a command to load the section groups with the specified properties.
    labsSectionGroups.load("id,name"); 
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index.
            if (labsSectionGroups.items.length > 0) {
                console.log("Section group name: " + labsSectionGroups.items[0].name);
                console.log("Section group ID: " + labsSectionGroups.items[0].id);
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


### <a name="getitem(index:-number-or-string)"></a>getItem(index: number or string)
Obtiene un grupo de secciones por id. o por su índice en la colección. Solo lectura.

#### <a name="syntax"></a>Sintaxis
```js
sectionGroupCollectionObject.getItem(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number o string|El id. del grupo de secciones, o bien la ubicación del índice del bloc de notas en la colección.|

#### <a name="returns"></a>Valores devueltos
[SectionGroup](sectiongroup.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtiene un grupo de secciones según su posición en la colección.

#### <a name="syntax"></a>Sintaxis
```js
sectionGroupCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[SectionGroup](sectiongroup.md)

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
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad

**items**
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups. 
    // For best performance, request specific properties.
    sectionGroups.load("name"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Iterate through the collection or access items individually by index, for example: sectionGroups.items[0]
            $.each(sectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);  
                console.log("Section group ID: " + sectionGroup.id);  
            });
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

