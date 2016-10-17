# <a name="officeextension.error-object-(javascript-api-for-onenote)"></a>Objeto OfficeExtension.Error (API de JavaScript para OneNote)

_Se aplica a: OneNote Online_


Representa los errores que se producen al usar la API de JavaScript para OneNote.

## <a name="properties"></a>Propiedades
| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|código|cadena|Obtiene un valor que indica el tipo de error. El valor puede ser "InvalidArgument", "GeneralException", "ItemNotFound" o "UnsupportedOperationForObjectType". |
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
