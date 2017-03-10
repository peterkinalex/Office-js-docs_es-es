# <a name="settingcollection-object-javascript-api-for-excel"></a>Objeto SettingCollection (API de JavaScript para Excel)

Representa una colección de objetos de hoja de cálculo que forman parte del libro.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|elementos|[Setting[]](setting.md)|Una colección de objetos de configuración. Solo lectura.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto Set|
|:---------------|:--------|:----------|:----|
|[add(key: string, value: (any)[])](#addkey-string-value-any)|[Setting](setting.md)|Establece o agrega la configuración especificada en el libro.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|entero|Obtiene el número de configuraciones de una colección.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|Obtiene una entrada de configuración mediante la clave.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Setting](setting.md)|Obtiene una entrada de configuración mediante la clave. Si el valor no existe, devolverá un objeto NULL.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="addkey-string-value-any"></a>add(key: string, value: (any)[])
Establece o agrega la configuración especificada en el libro.

#### <a name="syntax"></a>Sintaxis
```js
settingCollectionObject.add(key, value);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|Key|string|La clave de la nueva configuración.|
|valor|(any)[]|Valor de la nueva configuración.|

#### <a name="returns"></a>Valores devueltos
[Setting](setting.md)

### <a name="getcount"></a>getCount()
Obtiene el número de configuraciones de una colección.

#### <a name="syntax"></a>Sintaxis
```js
settingCollectionObject.getCount();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
int

### <a name="getitemkey-string"></a>getItem(key: string)
Obtiene una entrada de configuración mediante la clave.

#### <a name="syntax"></a>Sintaxis
```js
settingCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|Key|string|Clave de la configuración.|

#### <a name="returns"></a>Valores devueltos
[Setting](setting.md)

### <a name="getitemornullobjectkey-string"></a>getItemOrNullObject(key: string)
Obtiene una entrada de configuración mediante la clave. Si el valor no existe, devolverá un objeto NULL.

#### <a name="syntax"></a>Sintaxis
```js
settingCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>Parámetros
| Parámetro       | Tipo    |Descripción|
|:---------------|:--------|:----------|:---|
|Key|string|La clave de la configuración.|

#### <a name="returns"></a>Valores devueltos
[Setting](setting.md)
