
# Método Bindings.addFromNamedItemAsync
Agrega un enlace a un elemento con nombre del documento.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Modificado por última vez**|1.1|

```
Office.context.document.bindings.addFromNamedItemAsync(itemName, bindingType [, options], callback);
```


## Parámetros



|**Nombre**|**Tipo**|**Descripción**|**Notas de compatibilidad**|
|:-----|:-----|:-----|:-----|
| _itemName_|**string**|El nombre del elemento con nombre. Requerido.||
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|Especifica el tipo de objeto de enlace que se debe crear. Necesario. Devuelve **null** si el objeto seleccionado no se puede convertir en el tipo especificado.||
| _options_|**object**|Especifica cualquiera de los siguientes [parámetros opcionales](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _id_|**string**|Especifica el nombre único que se debe usar para identificar el nuevo objeto de enlace. Si no se pasa ningún argumento para el parámetro _id_, [Binding.id](../../reference/shared/binding.id.md) se genera automáticamente.||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** o **undefined**|Un elemento de cualquier tipo definido por el usuario que se devuelve en el objeto **AsyncResult** sin sufrir modificaciones.||
| _callback_|**object**|Una función que se invoca cuando se devuelve la devolución de llamada, cuyo único parámetro es del tipo **AsyncResult**.||

## Valor de devolución de llamada

Cuando la función que ha remitido al parámetro _callback_ se ejecute, recibirá un objeto [AsyncResult](../../reference/shared/asyncresult.md) al que puede obtener acceso desde el único parámetro de la función de devolución de llamada.

En la función de devolución de llamada que se ha remitido al método **addFromNamedItemAsync**, puede usar las propiedades del objeto **AsyncResult** para devolver la siguiente información.



|**Propiedad**|**Usar para...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Tener acceso al objeto [Binding](../../reference/shared/binding.md) que representa el elemento con nombre especificado.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determinar si la operación se ha completado correctamente o no.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Tener acceso a un objeto [Error](../../reference/shared/error.md) que proporcione información sobre el error si la operación no se ha llevado a cabo correctamente.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Tener acceso al valor o al **object** definidos por el usuario si ha remitido uno como parámetro _asyncContext_.|

## Observaciones

 **Para Excel**, el parámetro _itemName_ puede referirse a un rango con nombre o a una tabla.

De forma predeterminada, al agregar una tabla en Excel se asigna el nombre "Tabla1" para la primera tabla que agregue, "Tabla2" para la segunda y así sucesivamente. Para asignar un nombre significativo para una tabla en la interfaz de usuario de Excel, use la propiedad **Nombre de la tabla** de la pestaña **Herramientas de tabla | | Diseño** de la cinta.


 >**Nota** En Excel, cuando se especifica una tabla como elemento con nombre, debe asignarle un nombre completo e incluir el nombre de la hoja de cálculo en el nombre de la tabla según este formato: `"Sheet1!Table1"`

 **Para Word**, el parámetro _itemName_ hace referencia a la propiedad **Título** de un control de contenido **Texto enriquecido**. (No puede enlazar con controles de contenido distintos del control de contenido **Texto enriquecido**).

De forma predeterminada, un control de contenido no tiene ningún valor **Title** asignado. Para asignar un nombre significativo en la interfaz de usuario de Word, después de insertar un control de contenido **Texto enriquecido** desde el grupo **Controles** de la pestaña **Desarrollador** de la cinta, use el comando **Propiedades** del grupo **Controles** para mostrar el cuadro de diálogo **Propiedades del control de contenido**. A continuación, establezca la propiedad **Title** del control de contenido con el nombre al que desee hacer referencia desde su código.


 >**Nota** En Word, si hay varios controles de contenido **Texto enriquecido** con el mismo valor de propiedad (nombre) **Título** e intenta enlazar con uno de estos controles de contenido mediante este método (especificando su nombre como el parámetro _itemName_), la operación fallará.


## Ejemplo

En el ejemplo siguiente se agrega un enlace al elemento con nombre `myRange` en Excel como un enlace "matriz" y se asigna el [id](../../reference/shared/binding.id.md) del enlace como `myMatrix`.


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

En el ejemplo siguiente se agrega un enlace al elemento con nombre `Table1` en Excel como un enlace "tabla" y se asigna el **id** del enlace como `myTable`.




```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("Table1", "table", {id:'myTable'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

En el ejemplo siguiente se crea un enlace de texto en Word a un control de contenido de texto enriquecido llamado `"FirstName"`, se asigna el **id**`"firstName"` y, a continuación, se muestra esa información.




```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
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

||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|MatrixBindings, TableBindings, TextBindings|
|**Nivel de permisos mínimo**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|En los complementos para Excel, puede crear un enlace de tabla (remitiendo _bindingType_ como **Office.BindingType.Table**) en un rango de celdas que contenga datos tabulares, aunque no se hayan agregado a la hoja de cálculo como tabla (con los comandos **Insertar**  >  **Tablas**  > **Tabla** o **Inicio**  >  **Estilos**  >  **Dar formato como tabla**).|
|1.1|Se ha agregado compatibilidad para el enlace de tablas en los complementos de contenido para Access. |
|1.0|Agregado|

## Vea también



#### Otros recursos


[Enlazar a regiones de un documento u hoja de cálculo](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md#add-a-binding-to-a-named-item)
