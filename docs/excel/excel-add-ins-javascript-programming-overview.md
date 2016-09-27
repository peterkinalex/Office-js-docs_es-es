# Introducción a la programación de API de JavaScript para Excel

En este artículo se describe cómo usar la API de JavaScript para Excel para crear complementos de Excel 2016. Es una introducción a los conceptos clave que son fundamentales para usar las API, como RequestContext, los objetos proxy de JavaScript, sync(), Excel.run() y load(). Los ejemplos de código del final del artículo muestran cómo aplicar los conceptos.

## RequestContext

El objeto RequestContext facilita las solicitudes para la aplicación de Excel. Como el complemento de Office y la aplicación de Excel se ejecutan en dos procesos diferentes, se necesita un contexto de solicitud para tener acceso desde el complemento a Excel y a objetos relacionados, como hojas de cálculo y tablas. Un contexto de solicitud se crea como se muestra a continuación.

```js
var ctx = new Excel.RequestContext();
```

## Objetos proxy

Los objetos de JavaScript de Excel declarados y usados en un complemento son objetos proxy para los objetos reales de un documento de Excel. Las acciones llevadas a cabo en los objetos proxy no se realizan en Excel y el estado del documento de Excel no se realiza en los objetos proxy mientras no se sincronice el estado del documento. El estado del documento se sincroniza cuando se ejecuta context.sync() (véase a continuación).

Por ejemplo, el objeto de JavaScript local `selectedRange` se declara para que haga referencia al rango seleccionado. Esto puede usarse para poner en cola la configuración de sus propiedades y métodos de invocación. Las acciones en dichos objetos no se realizan hasta que se ejecuta el método sync().

```js
var selectedRange = ctx.workbook.getSelectedRange();
```

## sync()

El método sync() disponible en el contexto de solicitud sincroniza el estado entre los objetos proxy de JavaScript y los objetos reales de Excel. Para ello, ejecuta las instrucciones situadas en la cola en el contexto y recupera las propiedades de los objetos de Office cargados para usarlos en el código.  Este método devuelve una promesa, que se resuelve cuando se completa la sincronización.

## Excel.run(function(context) { batch })

Excel.run() ejecuta un script por lotes que realiza acciones en el modelo de objetos de Excel. Los comandos por lotes incluyen definiciones de objetos proxy locales de JavaScript y métodos sync() que sincronizan el estado entre los objetos locales y de Excel y la resolución de la promesa. La ventaja de procesamiento por lotes de las solicitudes en Excel.run() es que, cuando se resuelve la promesa, los objetos de intervalo de los que se realiza el seguimiento y que se asignaron durante la ejecución se liberarán automáticamente.

El método de ejecución toma RequestContext y devuelve una promesa que, normalmente, solo es el resultado de ctx.sync(). Es posible ejecutar la operación por lotes fuera de Excel.run(). Sin embargo, en este caso, todas las referencias a objetos de intervalo deben seguirse y administrarse manualmente.

## load()

El método load() se usa para rellenar los objetos proxy creados en la capa de JavaScript del complemento. Al intentar recuperar un objeto, como una hoja de cálculo, se crea en primer lugar un objeto proxy local en la capa de JavaScript. Dicho objeto puede usarse para poner en cola la configuración de sus propiedades y métodos de invocación. Sin embargo, para leer las propiedades o las relaciones de los objetos, deben invocarse primero los métodos load() y sync(). El método load() toma las propiedades y las relaciones que necesitan cargarse cuando se llama al método sync().

_Sintaxis:_

```js
object.load(string: properties);
//or
object.load(array: properties);
//or
object.load({loadOption});
```
Donde,

* `properties` es la lista de propiedades o nombres de relaciones que se van a cargar, especificados como cadenas delimitadas por comas o como una matriz de nombres. Consulte los métodos .load() de cada objeto para obtener más detalles.
* `loadOption` especifica un objeto que describe las opciones selection, expansion, top y skip. Consulte las [opciones](../../reference/excel/loadoption.md) de carga de objetos para obtener más detalles.

## Ejemplo: Escribir valores de una matriz a un objeto de intervalo

En el ejemplo siguiente se muestra cómo escribir los valores de una matriz a un objeto de intervalo.

Excel.run() contiene un lote de instrucciones. Como parte de este lote, se crea un objeto proxy que hace referencia a un intervalo (dirección A1:B2) de la hoja de cálculo activa. El valor de este objeto proxy de intervalo se establece localmente. Para poder leer los valores, se le indica a la propiedad `text` del rango que debe cargarse en el objeto proxy. Todos estos comandos se ponen en cola y se ejecutan cuando se llama a ctx.sync(). El método sync() devuelve una promesa que puede usarse para encadenarla con otras operaciones.

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) {

    // Create a proxy object for the sheet
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    // Values to be updated
    var values = [
                 ["Type", "Estimate"],
                 ["Transportation", 1670]
                 ];
    // Create a proxy object for the range
    var range = sheet.getRange("A1:B2");

    // Assign array value to the proxy object's values property.
    range.values = values;

    // Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
    return ctx.sync().then(function() {
            console.log("Done");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

## Ejemplo: Copiar valores

En el ejemplo siguiente se muestra cómo copiar los valores del intervalo A1:A2 a B1:B2 de la hoja de cálculo activa usando el método load() del objeto de intervalo.

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) {

    // Create a proxy object for the range
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");

    // Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
    return ctx.sync().then(function() {
        // Assign the previously loaded values to the new range proxy object. The values will be updated once the following .then() function is invoked.
        ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = range.values;
    });
}).then(function() {
      console.log("done");
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

## Selección de propiedades y relaciones

De forma predeterminada, object.load() selecciona todas las propiedades escalares y complejas del objeto que se está cargando. Las relaciones no se cargan de forma predeterminada (por ejemplo, el formato es un objeto de relación del objeto de intervalo). Sin embargo, recomendamos que marque las propiedades y las relaciones para que se carguen explícitamente y mejoren así el rendimiento. Para ello, especifique (en el parámetro `load()`) un subconjunto de propiedades y relaciones para que se incluya en la respuesta. El método Load permite dos tipos de entradas:

* Nombres de propiedad y de relación como nombres de cadenas separados por comas _o bien_ como una matriz de cadenas con los nombres de propiedad o relación.
* Un objeto que describe las opciones selection, expansion, top y skip. Consulte las [opciones](../../reference/excel/loadoption.md) de carga de objetos para obtener más detalles.

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

### Ejemplo

La instrucción Load siguiente carga todas las propiedades del intervalo y, a continuación, expande el formato y el formato/relleno.

```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:B2";
    var myRange = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load(["address", "format/*", "format/fill", "entireRow" ]);
    return ctx.sync().then(function() {
        console.log (myRange.address); //ok
        console.log (myRange.format.wrapText); //ok
        console.log (myRange.format.fill.color); //ok
        //console.log (myRange.format.font.color); //not ok as it was not loaded

    });
}).then(function() {
      console.log("done");
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

## Null-Input

### Entrada null en una matriz bidimensional

La entrada `null` dentro de una matriz bidimensional (para valores, formato numérico, fórmulas) se omite en la API de actualización. No se producirá ninguna actualización en el objetivo previsto cuando la entrada `null` se envíe en valores, en formato numérico o en una cuadrícula de fórmulas de valores.

Ejemplo: Para actualizar solamente partes específicas del rango, como el formato numérico de algunas celdas, y conservar el formato numérico existente en otras partes del rango, establezca el formato numérico deseado donde sea necesario y envíe `null` para las demás celdas.

En la solicitud de establecimiento siguiente, solo se establecen algunas partes del formato numérico de intervalo, mientras se conserva el formato numérico existente en la parte restante (pasando valores null).

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
### Entrada null para una propiedad

`null` no es una entrada única válida para toda la propiedad. Por ejemplo, lo que se muestra a continuación no es válido, ya que no es posible establecer todos los valores en null u omitirlos.

```js
 range.values= null;

```

Lo que se muestra a continuación tampoco es válido, ya que null no es un valor de color válido.

```js
 range.format.fill.color =  null;
```

### Respuesta null

La representación de propiedades de formato con valores no uniformes produciría la devolución de un valor null en la respuesta.

Ejemplo: Un intervalo puede consistir en una o más celdas. En los casos en los que las celdas individuales contenidas en el intervalo especificado no tienen valores de formato uniformes, la representación del nivel del intervalo será indefinida.

```js
  "size" : null,
  "color" : null,
```

### Entrada y salida en blanco

Los valores en blanco en las solicitudes de actualización se tratan como instrucciones para borrar o restablecer la propiedad correspondiente. Un valor en blanco se representa mediante dos comillas dobles sin espacio en medio. `""`

Ejemplo:

* En el caso de `values`, el valor del rango se borra. Esto equivale a borrar el contenido de la aplicación.

* En el caso de `numberFormat`, el formato numérico se establece en `General`.

* En cuanto a `formula` y `formulaLocale`, los valores de la fórmula se borran.


En las operaciones de lectura, lo que cabe esperar es recibir valores en blanco si el contenido de las celdas está en blanco. Si la celda no contiene datos ni valores, la API devuelve un valor en blanco. Un valor en blanco se representa mediante dos comillas dobles sin espacio en medio. `""`.

```js
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```js
  range.formula = [["", "", "=Rand()"]];
```

## Rango sin delimitar

### Lectura

Una dirección de intervalo sin delimitar solo contiene identificadores de columna o fila e identificadores de fila o columna no especificadas (respectivamente), tales como:

* `C:C`, `A:F`, `A:XFD` (contiene filas no especificadas)
* `2:2`, `1:4`, `1:1048546` (contiene columnas no especificadas)

Cuando la API realiza una solicitud para recuperar un rango sin delimitar (por ejemplo, `getRange('C:C')`), la respuesta devuelta contiene `null` para las propiedades de nivel de celda, como `values`, `text`, `numberFormat`, `formula`, etc. Otras propiedades de rango como `address`, `cellCount`, etc. reflejarán el rango sin delimitar.

### Escritura

El establecimiento de propiedades de nivel de celda (como valores, formato numérico, etc.) en el intervalo sin delimitar **no está permitido**, ya que la solicitud de entrada podría ser demasiado grande para controlarla.

Ejemplo: Lo que se muestra a continuación no es una solicitud de actualización válida porque el intervalo solicitado está sin delimitar.

```js
...
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A:B");
    range.values = 'Due Date';
...
```

Cuando se intenta realizar una operación de actualización en un intervalo así, la API devuelve un error.


## Intervalo grande

Un intervalo grande es aquel cuyo tamaño es demasiado grande para una sola llamada a la API. Numerosos factores contenidos en el intervalo, como el número de celdas, los valores, el formato numérico y las fórmulas, pueden hacer que la respuesta sea tan grande que resulte inadecuada para la interacción con la API. La API hace lo que puede para devolver o escribir en los datos solicitados. Sin embargo, el gran tamaño puede generar una condición de error de la API debido al elevado uso de recursos.

Para evitarlo, recomendamos usar la lectura o la escritura para el intervalo grande en varios tamaños de intervalo pequeños.


## Copia de entrada única

Para poder actualizar un intervalo con los mismos valores o formato numérico o aplicar la misma fórmula en un intervalo, se usa la siguiente convención en la API establecida. En Excel, este comportamiento es similar a la entrada de valores o fórmulas en un intervalo en el modo CTRL+ENTRAR.

La API buscará un *valor de celda único* y, si la dimensión del intervalo de destino no coincide con la dimensión del intervalo de entrada, aplicará la actualización a todo el intervalo en el modo CTRL+ENTRAR con el valor o la fórmula proporcionados en la solicitud.

### Ejemplos

La siguiente solicitud actualiza el intervalo seleccionado con el texto "Due Date". Tenga en cuenta que el intervalo tiene 20 celdas, mientras que el texto proporcionado solo tiene un valor de celda.

```js
Excel.run(function (ctx) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.values = 'Due Date';
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

La siguiente solicitud actualiza el intervalo seleccionado con la fecha "3/11/2015".

```js
Excel.run(function (ctx) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.numberFormat = 'm/d/yyyy';
    range.values = '3/11/2015';
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
La siguiente solicitud actualiza el intervalo seleccionado con una fórmula que se aplicará en todo el intervalo en el modo CTRL+ENTRAR.

```js
Excel.run(function (ctx) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.numberFormat = 'm/d/yyyy';
    range.values = '3/11/2015';
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


## Mensajes de error

Los errores se devuelven usando un objeto de error que consta de un código y un mensaje. En la siguiente tabla se proporciona una lista de las posibles condiciones de error que pueden producirse.

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |El argumento no es válido, o falta o tiene un formato incorrecto.|
|InvalidRequest  |No se puede procesar la solicitud.|
|InvalidReference|Esta referencia no es válida para la operación actual.|
|InvalidBinding  |Este enlace de objeto ya no es válido debido a actualizaciones anteriores.|
|InvalidSelection|La selección actual no es válida para esta operación.|
|Unauthenticated |La información de autenticación necesaria falta o no es válida.|
|AccessDenied   |No se puede realizar la operación solicitada.|
|ItemNotFound   |El recurso solicitado no existe.|
|ActivityLimitReached|Se alcanzó el límite de actividad.|
|GeneralException|Se produjo un error interno al procesar la solicitud.|
|NotImplemented  |La característica solicitada no se implementó.|
|ServiceNotAvailable|El servicio no está disponible.|
|Conflict   |No se pudo procesar la solicitud debido a un conflicto.|
|ItemAlreadyExists|El recurso que se está creando ya existe.|
|UnsupportedOperation|No se admite la operación que se está intentando.|
|RequestAborted|La solicitud se anuló durante el tiempo de ejecución.|
|ApiNotAvailable|La API solicitada no está disponible.|
|InsertDeleteConflict|La operación de inserción o eliminación intentada dio lugar a un conflicto.|
|InvalidOperation|La operación intentada no es válida en el objeto.|

## Recursos adicionales

* [Crear su primer complemento de Excel](build-your-first-excel-add-in.md)
* [Explorador de fragmentos de código](https://github.com/OfficeDev/office-js-snippet-explorer)
* [Ejemplos de código de complementos de Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Referencia de la API de JavaScript de complementos de Excel](excel-add-ins-javascript-api-reference.md)
