
# Método Bindings.addFromPromptAsync
 Muestra la interfaz de usuario que permite al usuario especificar una selección a la que enlazar.

|||
|:-----|:-----|
|**Hosts:**|Access y Excel|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|No en un conjunto|
|**Modificado por última vez**|1.1|

```
_bindingsObj.addFromPromptAsync(bindingType [, options], callback);
```


## Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|Especifica el tipo de objeto de enlace que se debe crear. Necesario. Devuelve **null** si el objeto seleccionado no se puede convertir en el tipo especificado.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _id_|**string**|Especifica el nombre único que se debe usar para identificar el nuevo objeto de enlace. Si no se pasa ningún argumento para el parámetro _id_, [Binding.id](../../reference/shared/binding.id.md) se genera automáticamente.||
| _promptText_|**string**|Especifica la cadena que se muestra en la interfaz de usuario del símbolo del sistema que informa al usuario de lo que debe seleccionar. Limitado a 200 caracteres. Si no se pasa ningún argumento _promptText_, se muestra "Haga una selección".||
| _sampleData_|[TableData](../../reference/shared/tabledata.md)|Especifica una tabla de datos de ejemplo que se muestra en la interfaz de usuario del símbolo del sistema como un ejemplo de los tipos de campos (columnas) que puede enlazar el complemento. Los encabezados que hay en el objeto **TableData** especifican las etiquetas usadas en la interfaz de usuario de la selección de campos. Opcional. **Nota:** Este parámetro solo se usa en los complementos para Access. Se ignora si se proporciona al llamar al método en un complemento para Excel.||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **addFromPromptAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Tener acceso al objeto [Binding](../../reference/shared/binding.md) que representa la selección especificada por el usuario.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **object** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## Comentarios

Agrega un objeto de enlace del tipo especificado a la colección [Bindings](../../reference/shared/bindings.bindings.md), que se identificará con el _id_ proporcionado. Si no se puede enlazar la selección especificada, el método falla.


## Ejemplo




```js
function addBindingFromPrompt() {

    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'MyBinding', promptText: 'Select text to bind to.' }, function (asyncResult) {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|No en un conjunto|
|**Nivel de permisos mínimo**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel en Office para iPad.|
|1.1|En los complementos para Excel, puede crear un enlace de tabla (remitiendo _bindingType_ como **Office.BindingType.Table**) en un rango de celdas que contenga datos tabulares, aunque no se hayan agregado a la hoja de cálculo como tabla en la interfaz de usuario de Excel (con los comandos **Insertar**  >  **Tablas**  > **Tabla** o **Inicio**  >  **Estilos**  >  **Dar formato como tabla**).|
|1.1|Se ha agregado compatibilidad para el enlace de tablas en los complementos de contenido para Access. |
|1.1|Se ha agregado compatibilidad para enlazar los datos de matriz como enlace de tabla en los complementos para Excel.|
|1.0|Agregado|
