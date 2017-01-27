# <a name="pagecollection-object-javascript-api-for-visio"></a>Objeto PageCollection (API de JavaScript para Visio)

Se aplica a: _Visio Online_
>**Nota:** Las API de JavaScript para Visio no están disponibles actualmente para su uso en entornos de producción o de versión preliminar.

Representa una colección de objetos Page que forman parte del documento.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|elementos|[Page[]](page.md)|Colección de objetos de página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-items)|

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|Obtiene el número de páginas de una colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-getCount)|
|[getItem(clave: número o cadena)](#getitemkey-number-or-string)|[Página](page.md)|Obtiene una página mediante su clave (nombre o identificador).|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-load)|

## <a name="method-details"></a>Detalles del método


### <a name="getcount"></a>getCount()
Obtiene el número de páginas de una colección.

#### <a name="syntax"></a>Sintaxis
```js
pageCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
int

### <a name="getitemkey-number-or-string"></a>getItem(clave: número o cadena)
Obtiene una página mediante su clave (nombre o identificador).

#### <a name="syntax"></a>Sintaxis
```js
pageCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|Key|número o cadena|Clave es el nombre o el identificador de la página que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[Page](page.md)

#### <a name="examples"></a>Ejemplos
```js
Visio.run(function (ctx) { 
    var pageName = 'Page-1';
    var page = ctx.document.pages.getItem(pageName);
    page.activate();
    return ctx.sync();
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
void
