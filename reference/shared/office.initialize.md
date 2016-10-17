
# <a name="office.initialize-event"></a>Evento Office.initialize
Se produce cuando se carga el entorno en tiempo de ejecución y el complemento está preparado para empezar a interactuar con la aplicación y el documento alojado. 

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project y Word|
|**Modificado por última vez en**|1.1|

```js
Office.initialize = function (reason) {/* initialization code */}
```


## <a name="remarks"></a>Comentarios

El parámetro _reason_ de la función de escucha del evento **initialize** devuelve un valor de la enumeración [InitializationReason](../../reference/shared/initializationreason-enumeration.md) que especifica cómo se ha producido la inicialización. El complemento de contenido o panel de tareas puede inicializarse con dos procedimientos:


- El usuario lo inserta desde la sección **Complementos usados recientemente** de la lista desplegable **Complemento** que aparece en la pestaña **Insertar** de la cinta de opciones de la aplicación host de Office o bien desde el cuadro de diálogo **Insertar complemento**.
    
- El usuario abre un documento que contiene el complemento.
    

 >**Nota**: el parámetro reason de la función de escucha del evento **initialize** solo devuelve un valor de la enumeración **InitializationReason** para los complementos de contenido y panel de tareas (para los complementos de Outlook no devuelve ningún valor).


## <a name="example"></a>Ejemplo

Puede usar el valor de **InitializationEnumeration** para implementar diferentes lógicas si se ha insertado el complemento por primera vez (es decir, si la aplicación no formaba parte del documento previamente). El ejemplo siguiente muestra un caso de lógica simple que usa el valor del parámetro _reason_ para mostrar cómo se ha inicializado el complemento de contenido o panel de tareas.


```js
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, code specific to the add-in can run.
    // Display initialization reason.
    if (reason == "inserted")
    write("The add-in was just inserted.");

    if (reason == "documentOpened")
    write("The add-in is already part of the document.");
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este evento es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este evento.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|**OWA para dispositivos**|**Outlook para Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||v||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Project**|v|||||
|**Word**|v|v|v|||

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, Outlook y panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint Online.|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para inicializar complementos de contenido para Access.|
|1.0|Agregado|
