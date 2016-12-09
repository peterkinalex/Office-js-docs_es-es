# <a name="settingcollection-object-javascript-api-for-excel"></a>Objeto SettingCollection (API de JavaScript para Excel)

Representa una colección de objetos de hoja de cálculo que forman parte del libro.

## <a name="properties"></a>Propiedades

| Propiedad     | Tipo   |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|items|[Setting[]](setting.md)|Una colección de objetos de configuración. Solo lectura.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|Obtiene una entrada de configuración mediante la clave.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: string)](#getitemornullkey-string)|[Setting](setting.md)|Obtiene una entrada de configuración mediante la clave. Si la configuración no existe, la propiedad isNull del objeto devuelto será True.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Rellena el objeto proxy que se ha creado en la capa de JavaScript con los valores de propiedad y objeto especificados en el parámetro.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[set(key: string, value: string)](#setkey-string-value-string)|[Setting](setting.md)|Establece o agrega la configuración especificada en el libro.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="getitemkey-string"></a>getItem(key: string)
Obtiene una entrada de configuración mediante la clave.

#### <a name="syntax"></a>Sintaxis
```js
settingCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|Key|string|Clave de la configuración.|

#### <a name="returns"></a>Valores devueltos
[Setting](setting.md)

### <a name="getitemornullkey-string"></a>getItemOrNull(key: string)
Obtiene una entrada de configuración mediante la clave. Si la configuración no existe, la propiedad isNull del objeto devuelto será True.

#### <a name="syntax"></a>Sintaxis
```js
settingCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|Key|string|La clave de la configuración.|

#### <a name="returns"></a>Valores devueltos
[Setting](setting.md)

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

### <a name="setkey-string-value-string"></a>set(key: string, value: string)
Establece o agrega la configuración especificada en el libro.

#### <a name="syntax"></a>Sintaxis
```js
settingCollectionObject.set(key, value);
```

#### <a name="parameters"></a>Parámetros
| Parámetro    | Tipo   |Descripción|
|:---------------|:--------|:----------|:---|
|Key|string|La clave de la nueva configuración.|
|value|string|El valor de la nueva configuración.|

#### <a name="returns"></a>Valores devueltos
[Setting](setting.md)
