
# Método ProjectDocument.getSelectedDataAsync
Obtiene de forma asincrónica el valor de texto de los datos que contiene la selección actual de una o varias celdas del diagrama de Gantt.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selección|
|**Agregado en**|1,0|

```
Office.context.document.getSelectedDataAsync(coercionType[, options][, callback]);
```


## Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)|El tipo de la estructura de datos que se debe devolver. Obligatorio.<br/>Project 2013 solo admite **Office.CoercionType.Text** o `"text"`.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|El formato que se debe usar para los valores de fecha o de número.<br/>Project 2013 ignora este parámetro y lo establece de forma interna en `unformatted`.||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|Especifica si se deben incluir solo los datos visibles o todos los datos. <br/>Project 2013 ignora este parámetro y lo establece de forma interna en  `all`.||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## Valor de devolución de llamada

Cuando la función _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el parámetro de la función de devolución de llamada.

Para el método **getSelectedDataAsync**, el objeto devuelto [AsyncResult](../../reference/shared/asyncresult.md) contiene las siguientes propiedades.


****


|**Nombre**|**Descripción**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Los datos que se han pasado en el parámetro opcional _asyncContext_, si se usó el parámetro.|
|[error](../../reference/shared/asyncresult.error.md)|Información sobre el error si la propiedad **status** es igual a **failed**.|
|[estado](../../reference/shared/asyncresult.status.md)|Estado **succeeded** o **failed** de la llamada asincrónica.|
|[value](../../reference/shared/asyncresult.value.md)|El valor de texto de las celdas seleccionadas.|

## Comentarios

El método **ProjectDocument.getSelectedDataAsync** reemplaza el método [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) y devuelve el valor de texto de los datos conforme a la selección de una o varias celdas en la vista de diagrama de Gantt. **ProjectDocument.getSelectedDataAsync** admite solo formato de texto como [CoercionType](../../reference/shared/coerciontype-enumeration.md) (no admite `matrix`, `table` ni otros formatos).


## Ejemplo

En el ejemplo de código siguiente se obtienen los valores de las celdas seleccionadas. Usa el parámetro opcional _asyncContext_ para pasar texto a la función de devolución de llamada.

En el ejemplo se asume que el complemento tiene una referencia a la biblioteca de jQuery y que los controles de la siguiente página se definen en el div de contenido del cuerpo de la página.




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            $('#get-info').click(getSelectedText);
        });
    };

    // Get the text from the selected cells in the document, and display it in the add-in.
    function getSelectedText() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            {asyncContext: 'Some related info'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'Selected text: {0}<br/>Passed info: {1}',
                        result.value, result.asyncContext);
                    $('#message').html(output);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|
|:-----|:-----|:-----|
|**Project**|v||

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|Selección|
|**Nivel de permisos mínimo**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1,0|Agregado|

## Vea también



#### Otros recursos


[Objeto AsyncResult](../../reference/shared/asyncresult.md)

[Office.CoercionType](../../reference/shared/coerciontype-enumeration.md)

[Objeto ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
