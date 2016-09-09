# Objeto OfficeExtension.Error (API de JavaScript para Excel)

Representa los errores que se producen al usar la API de JavaScript para Excel.

_Se aplica a: Excel 2016, Excel Online, Excel para iOS y Office 2016_

## Propiedades
| Propiedad     | Tipo   |Descripción
|:---------------|:--------|:----------|
|código|cadena|Obtiene un valor que indica el tipo de error. El valor puede ser "AccessDenied", "ActivityLimitReached", "BadPassword", "GeneralException", "InsertDeleteConflict", "InvalidArgument", "InvalidBinding", "InvalidOperation", "InvalidReference", "InvalidSelection", "ItemAlreadyExists", "ItemNotFound", "NotImplemented" o "UnsupportedOperation". |
|debugInfo|string|Obtiene un valor que indica lo sucedido al producirse el error. Este valor solo está diseñado para su uso durante el desarrollo y la depuración.  |
|mensaje |cadena| Obtiene una cadena legible y localizada manualmente que corresponde al código de error.|
|nombre |cadena| Obtiene un valor que siempre es "OfficeExtension.Error". |
|traceMessages |string[]| Obtiene una matriz de valores que corresponden a los mensajes de instrumentación establecidos con context.trace();. |

## Métodos

| Método           | Tipo de valor devuelto    |Descripción|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|Devuelve el código de error y los valores del mensaje en el formato siguiente: "{0}: {1}", código, mensaje.|

## Detalles del método

### toString()
Devuelve el código de error y los valores del mensaje en el formato siguiente: "{0}: {1}", código, mensaje.

#### Sintaxis
```js
error.toString()
```

#### Parámetros
Ninguno

#### Valores devueltos
cadena
