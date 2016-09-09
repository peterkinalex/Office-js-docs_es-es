

# Método Office.select
Crea una promesa de devolver un enlace basado en la cadena de selector transferida.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Disponible en [Conjuntos de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**Modificado por última vez en**|1.1|

```js
Office.select(str, onError);
```


## Parámetros


_str_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **string**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;La cadena de selector que se quiere analizar y para la que debe crearse una promesa.

_onError_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **function**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**. Opcional.
    

## Valor de devolución de llamada

Cuando se ejecute la función que ha enviado al parámetro _onError_, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el parámetro único de la función de devolución de llamada. Si la operación falla, use la propiedad [AsyncResult.error](../../reference/shared/asyncresult.error.md) para obtener acceso a un objeto [Error](../../reference/shared/error.md) con información sobre el error.


## Comentarios

El método **Office.select** proporciona acceso a una promesa de objeto [Binding](../../reference/shared/binding.md) que intenta devolver el enlace especificado cuando se invoca cualquiera de sus métodos asincrónicos.

Formatos compatibles: "bindings# _bindingId_", que devuelve un objeto **Binding** para el enlace con el [identificador](../../reference/shared/binding.id.md) de `bindingId`. Para obtener más información, consulte [Programación asincrónica en complementos de Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md#asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings) y [Enlazar a regiones de un documento u hoja de cálculo](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


 >**Nota**: Si la promesa del método **select** devuelve correctamente un objeto **Binding**, dicho objeto solo expone los siguientes cuatro métodos del objeto [Binding](../../reference/shared/binding.md): [getDataAsync](../../reference/shared/binding.getdataasync.md), [setDataAsync](../../reference/shared/binding.setdataasync.md), [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) y [removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md). Si la promesa no puede devolver un objeto **Binding**, se puede usar la devolución de llamada _onError_ para obtener acceso a un objeto [asyncResult.error](../../reference/shared/asyncresult.error.md) y obtener más información. Si necesita llamar a un miembro del objeto **Binding** que no sea ninguno de los cuatro métodos expuestos por la promesa del objeto **Binding** devuelta por el método **select**, use en su lugar el método [getByIdAsync](../../reference/shared/bindings.getbyidasync.md) con la propiedad [Document.bindings](../../reference/shared/document.bindings.md) y el método [Bindings.getByIdAsync](../../reference/shared/bindings.getbyidasync.md) para recuperar el objeto **Binding**.


## Ejemplo

En el ejemplo de código siguiente se usa el método **select** para recuperar un enlace con el **id.** " `cities`" de la colección de **Bindings** y, después, se realiza una llamada al método [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) para agregar un controlador de eventos para el evento [dataChanged](../../reference/shared/binding.bindingdatachangedevent.md) del enlace.


```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}
```




## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**Nivel de permisos mínimo**|[ReadDocument (ReadAllDocument para Open Office XML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Se ha agregado el uso del método **select** para devolver enlaces de tabla creados en complementos de contenido para Access.|
|1.0|Agregado|
