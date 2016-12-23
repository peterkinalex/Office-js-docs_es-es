# <a name="shapedataitemcollection-object-javascript-api-for-visio"></a>Objeto ShapeDataItemCollection (API de JavaScript para Visio)

Se aplica a: _Visio Online_
>**Nota:** Las API de JavaScript para Visio están actualmente en la vista previa y están sujetas a cambios. Las API de JavaScript para Visio no se admiten actualmente para su uso en entornos de producción.

Representa el ShapeDataItemCollection de una forma determinada.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|elementos|[ShapeDataItem[]](shapedataitem.md)|Colección de objetos shapeDataItem. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-items)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|Obtiene el número de elementos de datos de formas.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-getCount)|
|[getItem(clave: cadena)](#getitemkey-string)|[ShapeDataItem](shapedataitem.md)|Obtiene el ShapeDataItem mediante su nombre.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-load)|

## <a name="method-details"></a>Detalles del método


### <a name="getcount"></a>getCount()
Obtiene el número de elementos de datos de formas.

#### <a name="syntax"></a>Sintaxis
```js
shapeDataItemCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
int

### <a name="getitemkey-string"></a>getItem(clave: cadena)
Obtiene el ShapeDataItem mediante su nombre.

#### <a name="syntax"></a>Sintaxis
```js
shapeDataItemCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|Key|string|La clave es el nombre de la ShapeDataItem que se va a recuperar.|

#### <a name="returns"></a>Valores devueltos
[ShapeDataItem](shapedataitem.md)

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
    var shape = activePage.shapes.getItem(0);
        var shapeDataItems = shape.shapeDataItems;
        shapeDataItems.load();
        return ctx.sync().then(function() {
            for (var i = 0; i < shapeDataItems.items.length; i++)
            {
                console.log(shapeDataItems.items[i].label);
                console.log(shapeDataItems.items[i].value);
            }
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
