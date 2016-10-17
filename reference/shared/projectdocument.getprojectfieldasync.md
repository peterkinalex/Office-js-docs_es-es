
# <a name="projectdocument.getprojectfieldasync-method"></a>Método ProjectDocument.getProjectFieldAsync
Obtiene de forma asincrónica el valor del campo especificado del proyecto activo.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selección|
|**Agregado en**|1.0|

```
Office.context.document.getProjectFieldAsync(fieldId[, options][, callback]);
```


## <a name="parameters"></a>Parámetros



|**Nombre**|**Tipo**|**Descripción**|
|:-----|:-----|:-----|:-----|
| _fieldId_|[ProjectProjectFields](../../reference/shared/projectprojectfields-enumeration.md)|El ID del campo de destino. Obligatorio.|
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):|
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.|
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.|

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el parámetro de la función de devolución de llamada.

Para el método **getProjectFieldAsync**, el objeto devuelto [AsyncResult](../../reference/shared/asyncresult.md) contiene las siguientes propiedades.


****


|**Nombre**|**Descripción**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Datos pasados en el parámetro opcional _asyncContext_, si se usó el parámetro.|
|[error](../../reference/shared/asyncresult.error.md)|Información sobre el error si la propiedad **status** es igual a **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Estado **succeeded** o **failed** de la llamada asincrónica.|
|[value](../../reference/shared/asyncresult.value.md)|Contiene la propiedad **fieldValue**, que representa el valor del campo especificado.|

## <a name="example"></a>Ejemplo

En el ejemplo de código siguiente se obtienen los valores de tres campos especificados para el proyecto activo y, a continuación, se muestran los valores en el complemento.

El ejemplo llama a **getProjectFieldAsync** recursivamente, después de que la llamada anterior se devuelva correctamente. También realiza un seguimiento de las llamadas para determinar cuándo se envían todas las llamadas.

En el ejemplo se presupone que el complemento tiene una referencia a la biblioteca de jQuery y que el siguiente control de página se define en div de contenido en el cuerpo de la página.




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Get information for the active project.
            getProjectInformation();
        });
    };

    // Get the specified fields for the active project.
    function getProjectInformation() {
        var fields =
            [Office.ProjectProjectFields.Start, Office.ProjectProjectFields.Finish, Office.ProjectProjectFields.GUID];
        var fieldValues = ['Start: ', 'Finish: ', 'GUID: '];
        var index = 0; 
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == fields.length) {
                var output = '';
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }
            else {
                Office.context.document.getProjectFieldAsync(
                    fields[index],
                    function (result) {

                        // If the call is successful, get the field value and then get the next field.
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
|**Disponible en los conjuntos de requisitos**||
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


[Enumeración ProjectProjectFields](../../reference/shared/projectprojectfields-enumeration.md)

[Objeto AsyncResult](../../reference/shared/asyncresult.md)

[Objeto ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
