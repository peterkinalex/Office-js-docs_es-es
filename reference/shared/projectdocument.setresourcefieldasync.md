

# <a name="projectdocument.setresourcefieldasync-method"></a>Método ProjectDocument.setResourceFieldAsync
Establece de forma asincrónica el valor del campo especificado para el recurso concreto.  **Importante:** Esta API solo funciona en Project 2016 para escritorio de Windows.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selección|
|**Agregado en**|1.1|

```js
Office.context.document.setResourceFieldAsync(resourceId, fieldId, fieldValue[, options][, callback]);
```


## <a name="parameters"></a>Parámetros

_resourceId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;El GUID del recurso. Obligatorio.
    
_fieldId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Identificador del campo de destino, ya sea como una constante [ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md) o como el valor entero correspondiente. Obligatorio.
    
_fieldValue_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Valor para el campo de destino, de tipo **string**, **number**, **boolean** u **object**. Obligatorio.
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;El [parámetro opcional](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) siguiente:

&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **array, boolean, null, number, object, string** o **undefined**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto [AsyncResult](../../reference/shared/asyncresult.md) sin sufrir modificaciones. Opcional.</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Por ejemplo, puede pasar el argumento _asyncContext_ usando el formato `{asyncContext: 'Some text'}` o `{asyncContext: <object>}`.


_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **function**

&nbsp;&nbsp;&nbsp;&nbsp;Una función que se invoca cuando se devuelve la llamada al método, cuyo único parámetro es del tipo [AsyncResult](../../reference/shared/asyncresult.md). Opcional.

    

## <a name="callback-value"></a>Valor de devolución de llamada

Cuando la función _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el parámetro de la función de devolución de llamada.

En el caso del método **setResourceFieldAsync**, el objeto [AsyncResult](../../reference/shared/asyncresult.md) devuelto contiene las siguientes propiedades.


|**Nombre**|**Descripción**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Datos pasados en el parámetro opcional _asyncContext_, si se usó el parámetro.|
|[error](../../reference/shared/asyncresult.error.md)|Información sobre el error si la propiedad **status** es igual a **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Estado **succeeded** o **failed** de la llamada asincrónica.|
|[value](../../reference/shared/asyncresult.value.md)|Este método no devuelve ningún valor.|

## <a name="remarks"></a>Comentarios

Primero llame al método [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) o [getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md) para obtener el GUID del recurso y luego pase el GUID como el argumento _resourceId_ a **setResourceFieldAsync**. Solo se puede actualizar un único campo para un único recurso en cada llamada asincrónica.


## <a name="example"></a>Ejemplo

En el siguiente ejemplo de código se llama a [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) para obtener el GUID del recurso seleccionado en ese momento en una vista de recursos. Luego se establecen dos valores de campo de recursos al llamar a **setResourceFieldAsync** recursivamente.

El método [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) usado en el ejemplo exige que la vista activa sea una vista de tareas (por ejemplo, Uso de tareas) y que haya una tarea seleccionada. Consulte el método [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) para ver un ejemplo que activa un botón en función del tipo de vista activa.

En el ejemplo se asume que el complemento tiene una referencia a la biblioteca de jQuery y que los controles de la siguiente página se definen en el div de contenido del cuerpo de la página.




```HTML
<input id="set-info" type="button" value="Set info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#set-info').click(setResourceInfo);
        });
    };

    // Get the GUID of the resource, and then get the resource fields.
    function setResourceInfo() {
        getResourceGuid().then(
            function (data) {
                setResourceFields(data);
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

    // Set the specified fields for the selected resource.
    function setResourceFields(resourceGuid) {
        var targetFields = [Office.ProjectResourceFields.StandardRate, Office.ProjectResourceFields.Notes];
        var fieldValues = [.28, 'Notes for the resource.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setResourceFieldAsync(
                resourceGuid,
                targetFields[i],
                fieldValues[i],
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        i++;
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
        $('#message').html('Field values set');
    }

    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
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
|**Nivel de permisos mínimo**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad

|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Agregado|

## <a name="see-also"></a>Vea también



#### <a name="other-resources"></a>Otros recursos


[getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)
[getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md)
[Objeto AsyncResult](../../reference/shared/asyncresult.md)
[Enumeración ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md)
[Objeto ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)

