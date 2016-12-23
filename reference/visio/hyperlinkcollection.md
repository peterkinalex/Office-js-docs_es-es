# <a name="hyperlinkcollection-object-javascript-api-for-visio"></a>Objeto HyperlinkCollection (API de JavaScript para Visio)

Se aplica a: _Visio Online_
>**Nota:** Las API de JavaScript para Visio están actualmente en la vista previa y están sujetas a cambios. Las API de JavaScript para Visio no se admiten actualmente para su uso en entornos de producción.

Representa la colección del hipervínculo.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|elementos|[Hipervínculo[]](hyperlink.md)|Una colección de objetos de hipervínculo. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-items)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|Obtiene el número de hipervínculos.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-getCount)|
|[getItem(Clave: número o cadena)](#getitemkey-number-or-string)|[Hipervínculo](hyperlink.md)|Obtiene un hipervínculo mediante su clave (nombre o identificador).|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-load)|

## <a name="method-details"></a>Detalles del método


### <a name="getcount"></a>getCount()
Obtiene el número de hipervínculos.

#### <a name="syntax"></a>Sintaxis
```js
hyperlinkCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
int

### <a name="getitemkey-number-or-string"></a>getItem(Clave: número o cadena)
Obtiene un hipervínculo mediante su clave (nombre o identificador).

#### <a name="syntax"></a>Sintaxis
```js
hyperlinkCollectionObject.getItem(Key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|Clave|número o cadena|Clave es el nombre o el índice del hipervínculo que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[Hipervínculo](hyperlink.md)

### <a name="loadparam-object"></a>load(param: objeto)
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
### <a name="property-access-examples"></a>Ejemplos de acceso a la propiedad
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Manager Belt";
    var shape = activePage.shapes.getItem(shapeName);
    var hyperlinks = shape.hyperlinks;
    shapeHyperlinks.load();
        ctx.sync().then(function () {
            for(var i=0; i<shapeHyperlinks.items.length;i++)
                {
                  var hyperlink = shapeHyperlinks.items[i];
                  console.log("Description:"+hyperlink.description +"Address:"+hyperlink.address +"SubAddress:  "+ hyperlink.subAddress);
                }

            });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
