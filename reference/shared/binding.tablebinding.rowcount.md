
# <a name="tablebinding.rowcount-property"></a>Propiedad TableBinding.rowCount
Obtiene el número de filas de la tabla como un valor entero.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Modificado por última vez en Selección**|1.1|

```
var rowCount = bindingObj.rowCount;
```


## <a name="return-value"></a>Valor devuelto

El número de filas existentes en el objeto [TableBinding](../../reference/shared/binding.tablebinding.md) que se ha especificado.


## <a name="remarks"></a>Comentarios

Cuando se inserta una tabla vacía seleccionando una única fila en Excel 2013 y Excel Online (mediante **Tabla** en la pestaña **Insertar**), ambas aplicaciones host de Office crean una única fila de encabezados seguida de una única fila vacía. Sin embargo, si el script del complemento crea un enlace para esta tabla recién insertada (por ejemplo, con el método [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)) y, a continuación, comprueba el valor de la propiedad **rowCount**, el valor devuelto será distinto en función de si la hoja de cálculo se abre en Excel 2013 o en Excel Online.


- En Excel para el escritorio, **rowCount** devolverá 0 (no se cuenta la fila vacía que sigue a los encabezados).
    
- En Excel Online, **rowCount** devolverá 1 (se cuenta la fila vacía que sigue a los encabezados).
    
Para solucionar esta diferencia en el script, puede comprobar si `rowCount == 1` y, en caso afirmativo, comprobar si la fila contiene todas las cadenas vacías.

En el caso de los complementos de contenido para Access, la propiedad **rowCount** siempre devuelve -1 por motivos de rendimiento.


## <a name="example"></a>Ejemplo




```js
function showBindingRowCount() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Rows: " + asyncResult.value.rowCount);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|TableBindings|
|**Nivel de permisos mínimo**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para los complementos para Access.|
|1.0|Agregado|
