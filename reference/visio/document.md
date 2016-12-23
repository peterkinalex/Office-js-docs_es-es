# <a name="document-object-javascript-api-for-visio"></a>Objeto Document (API de JavaScript para Visio)

Se aplica a: _Visio Online_
>**Nota:** Las API de JavaScript para Visio están actualmente en la vista previa y están sujetas a cambios. Las API de JavaScript para Visio no se admiten actualmente para su uso en entornos de producción.

Representa la clase de documento.

## <a name="properties"></a>Propiedades

Ninguno

## <a name="relationships"></a>Relaciones
| Relación | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|aplicación|[Aplicación](application.md)|Representa una instancia de aplicación de Visio que contiene este documento. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-application)|
|pages|[PageCollection](pagecollection.md)|Representa una colección de páginas asociadas con el documento. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-pages)|

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|[getActivePage()](#getactivepage)|[Página](page.md)|Devuelve la página activa del documento.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-getActivePage)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-load)|
|[setActivePage(PageName: cadena)](#setactivepagepagename-string)|vacío|Establece la página activa del documento.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-setActivePage)|

## <a name="method-details"></a>Detalles del método


### <a name="getactivepage"></a>getActivePage()
Devuelve la página activa del documento.

#### <a name="syntax"></a>Sintaxis
```js
documentObject.getActivePage();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
[Page](page.md)

#### <a name="examples"></a>Ejemplos
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    var activePage = document.getActivePage();
    activePage.load();
    return ctx.sync().then(function () {
    console.log("pageName: " +activePage.name);
      });   
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
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
|:---------------|:--------|:----------|:---|
|param|object|Opcional. Acepta nombres de parámetro y de relación como una cadena delimitada o una matriz. O bien, proporciona el objeto [loadOption](loadoption.md).|

#### <a name="returns"></a>Valores devueltos
vacío

### <a name="setactivepagepagename-string"></a>setActivePage(PageName: cadena)
Establece la página activa del documento.

#### <a name="syntax"></a>Sintaxis
```js
documentObject.setActivePage(PageName);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|PageName|string|Nombre de la página|

#### <a name="returns"></a>Valores devueltos
void

#### <a name="examples"></a>Ejemplos
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    var pageName = "Page-1";
    document.setActivePage(pageName);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
```js
Visio.run(function (ctx) { 
    var pages = ctx.document.pages;
    var pageCount = pages.getCount();
    return ctx.sync().then(function () {
        console.log("Pages Count: " +pageCount.value);
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
```js
Visio.run(function (ctx) { 
    var application = ctx.document.application;
    application.showToolbars = false;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

