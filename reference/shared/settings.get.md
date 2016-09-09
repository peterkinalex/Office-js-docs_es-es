
# Método Settings.get
Recupera la configuración especificada.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Configuración|
|**Modificado por última vez en**|1.1|

```js
var mySetting = Office.context.document.settings.get(name);
```


## Parámetros



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **string**

&nbsp;&nbsp;&nbsp;&nbsp;Nombre, con distinción de mayúsculas y minúsculas, de la configuración que se debe recuperar.

    



## Valor devuelto

Un **object** que tiene nombres de propiedad asignados a valores serializados JSON.


## Ejemplo




```js
function displayMySetting() {
    write('Current value for mySetting: ' + Office.context.document.settings.get('mySetting'));
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|Configuración|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint Online.|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para crear configuraciones en los complementos de contenido para Access.|
|1.0|Agregado|
