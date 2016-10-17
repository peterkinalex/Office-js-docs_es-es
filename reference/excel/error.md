# <a name="officeextension.error-object-(javascript-api-for-excel)"></a>Objeto OfficeExtension.Error (API de JavaScript para Excel)

Representa los errores que se producen al usar la API de JavaScript para Excel.

_Se aplica a: Excel 2016, Excel Online, Excel para iOS, Office 2016_

## <a name="properties"></a>Propiedades
| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|código|cadena|Obtiene un valor que indica el tipo de error. El valor puede ser "AccessDenied", "ActivityLimitReached", "BadPassword", "GeneralException", "InsertDeleteConflict", "InvalidArgument", "InvalidBinding", "InvalidOperation", "InvalidReference", "InvalidSelection", "ItemAlreadyExists", "ItemNotFound", "NotImplemented" o "UnsupportedOperation". |
|debugInfo|string|Obtiene un valor que indica lo sucedido al producirse el error. Este valor solo está diseñado para su uso durante el desarrollo y la depuración.  |
|mensaje |cadena| Obtiene una cadena legible y localizada manualmente que corresponde al código de error.|
|nombre |cadena| Obtiene un valor que siempre es "OfficeExtension.Error". |
|traceMessages |string[]| Obtiene una matriz de valores que corresponden a los mensajes de instrumentación establecidos con context.trace();. |

## <a name="methods"></a>Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|Devuelve el código de error y los valores del mensaje en el formato siguiente: "{0}: {1}", código, mensaje.|

## <a name="method-details"></a>Detalles del método

### <a name="tostring()"></a>toString()
Devuelve el código de error y los valores del mensaje en el formato siguiente: "{0}: {1}", código, mensaje.

#### <a name="syntax"></a>Sintaxis
```js
error.toString()
```

#### <a name="parameters"></a>Parámetros
Ninguno

#### <a name="returns"></a>Valores devueltos
cadena
