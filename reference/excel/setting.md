# <a name="setting-object-javascript-api-for-excel"></a>Objeto Setting (API de JavaScript para Excel)

Setting representa un par clave-valor de una configuración que se conserva en el documento.

## <a name="properties"></a>Propiedades

| Propiedad       | Tipo    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|Key|string|Devuelve la clave que representa el identificador de la configuración. Solo lectura.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|valor|objeto|Representa el valor almacenado para esta configuración.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_Consulte los [ejemplos](#property-access-examples) de acceso a la propiedad._

## <a name="relationships"></a>Relaciones
Ninguno


## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción| Conjunto req.|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Elimina la configuración.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Detalles del método


### <a name="delete"></a>delete()
Elimina la configuración.

#### <a name="syntax"></a>Sintaxis
```js
settingObject.delete();
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
void
