
# <a name="projectdocument.getresourcefieldasync-method"></a>Método ProjectDocument.getResourceFieldAsync
Obtiene de forma asincrónica el valor del campo especificado para el recurso especificado en una vista de recursos.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selección|
|**Agregado en**|1.0|

```
Office.context.document.getResourceFieldAsync(resourceId, fieldId[, options][, callback]);
```


## <a name="parameters"></a>Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _resourceId_|**string**|El GUID del recurso. Obligatorio.||
| _fieldId_|[ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md)|El ID del campo de destino. Obligatorio.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el parámetro de la función de devolución de llamada.

Para el método **getResourceFieldAsync**, el objeto devuelto [AsyncResult](../../reference/shared/asyncresult.md) contiene las siguientes propiedades.


****


|**Nombre**|**Descripción**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Datos pasados en el parámetro opcional _asyncContext_, si se usó el parámetro.|
|[error](../../reference/shared/asyncresult.error.md)|Información sobre el error si la propiedad **status** es igual a **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Estado **succeeded** o **failed** de la llamada asincrónica.|
|[value](../../reference/shared/asyncresult.value.md)|Contiene la propiedad **fieldValue**, que representa el valor del campo especificado.|

## <a name="remarks"></a>Comentarios

Primero llame al método [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) para obtener el GUID del recurso, y luego páselo como el argumento _resourceId_ a **getResourceFieldAsync**. Si la vista activa no es una vista de recursos (por ejemplo, una vista de Uso de recursos u Hoja de recursos), o si no hay tareas seleccionadas en una vista de recursos, [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) devuelve un error 5001 (Error interno). Consulte [Método addHandlerAsync ](../../reference/shared/projectdocument.addhandlerasync.md) para ver un ejemplo que use el evento [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) y el método [getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md) para activar un botón según el tipo de vista activo.


## <a name="example"></a>Ejemplo

El siguiente código de ejemplo llama a [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) para obtener el GUID del recurso que está seleccionado en una vista de recursos. A continuación, obtiene tres valores de campo de recursos mediante una llamada a **getResourceFieldAsync** recursivamente.

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
            $('#get-info').click(getResourceInfo);
        });
    };

    // Get the GUID of the resource and then get the resource fields.
    function getResourceInfo() {
        getResourceGuid().then(
            function (data) {
                getResourceFields(data);
            }
        );
    }

    // Get the GUID of the selected resource.
    function getResourceGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Get the specified fields for the selected resource.
    function getResourceFields(resourceGuid) {
        var targetFields =
            [Office.ProjectResourceFields.Name, Office.ProjectResourceFields.Units, Office.ProjectResourceFields.BaseCalendar];
        var fieldValues = ['Name: ', 'Units: ', 'Base calendar: '];
        var index = 0; 
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == targetFields.length) {
                var output = '';
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }

            // If the call is successful, get the field value and then get the next field.
            else {
                Office.context.document.getResourceFieldAsync(
                    resourceGuid,
                    targetFields[index],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            fieldValues[index] += result.value.fieldValue;
                            getField(index++);
                        }
                        else {
                            onError(result.error);
                        }
                    }
                );
            }
        }
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|
|:-----|:-----|:-----|
|**Project**|v||

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|Selección|
|**Nivel de permisos mínimo**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.0|Agregado|

## <a name="see-also"></a>Vea también



#### <a name="other-resources"></a>Otros recursos


[Método getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedresourceasync.md)

[Enumeración ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md)

[Objeto AsyncResult](../../reference/shared/asyncresult.md)

[Objeto ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
