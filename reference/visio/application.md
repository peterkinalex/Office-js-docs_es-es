# <a name="application-object-javascript-api-for-visio"></a>Objeto Application (API de JavaScript para Visio)

Se aplica a: _Visio Online_
>**Nota:** Las API de JavaScript para Visio están actualmente en la vista previa y están sujetas a cambios. Las API de JavaScript para Visio no se admiten actualmente para su uso en entornos de producción.

Representa la aplicación.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|showToolbars|bool|Muestra u oculta las barras de herramientas estándares.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-application-showToolbars)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:---|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-application-load)|

## <a name="method-details"></a>Detalles del método


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