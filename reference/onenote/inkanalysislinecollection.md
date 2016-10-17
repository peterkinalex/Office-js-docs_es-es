# <a name="inkanalysislinecollection-object-(javascript-api-for-onenote)"></a>Objeto InkAnalysisLineCollection (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_  


Representa una colección de objetos InkAnalysisLine.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción|Comentarios|
|:---------------|:--------|:----------|:-------|
|count|int|Devuelve el número de InkAnalysisLines de la página. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-count)|
|items|[InkAnalysisLine[]](inkanalysisline.md)|Una colección de objetos InkAnalysisLine. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-items)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Comentarios|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkAnalysisLine](inkanalysisline.md)|Obtiene un objeto InkAnalysisLine por su identificador o por su índice en la colección. Solo lectura.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisLine](inkanalysisline.md)|Obtiene un objeto InkAnalysisLine en su posición en la colección.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[Ir](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-load)|

## <a name="method-details"></a>Detalles del método


### <a name="getitem(index:-number-or-string)"></a>getItem(index: number or string)
Obtiene un objeto InkAnalysisLine por su identificador o por su índice en la colección. Solo lectura.

#### <a name="syntax"></a>Sintaxis
```js
inkAnalysisLineCollectionObject.getItem(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number or string|El identificador del objeto InkAnalysisLine, o bien la ubicación del objeto InkAnalysisLine en la colección.|

#### <a name="returns"></a>Valores devueltos
[InkAnalysisLine](inkanalysisline.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtiene un objeto InkAnalysisLine en su posición en la colección.

#### <a name="syntax"></a>Sintaxis
```js
inkAnalysisLineCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|
|index|number|Valor de índice del objeto que se va a recuperar. Indizado con cero.|

#### <a name="returns"></a>Valores devueltos
[InkAnalysisLine](inkanalysisline.md)

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
