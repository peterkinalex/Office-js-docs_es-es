

# Método Office.useShortNamespace
Activa y desactiva el alias de `Office` para el espacio de nombres completo `Microsoft.Office.WebExtension`.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project y Word|
|**Modificado por última vez en**|1.1|

```js
Office.useShortNamespace(useShortcut);
```


## Parámetros



_useShortcut_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Tipo: **boolean**

    
&nbsp;&nbsp;&nbsp;&nbsp;**true** para usar alias de acceso directo; en caso contrario, **false** para deshabilitarlo. El valor predeterminado es **true**.
    


## Ejemplo



```js
function startUsingShortNamespace() {
    if (typeof Office === 'undefined') {
        Microsoft.Office.WebExtension.useShortNamespace(true);
    }
    else {
        Office.useShortNamespace(true);
    }
    write('Office alias is now ' + typeof Office);
}

function stopUsingShortNamespace() {
    if (typeof Office === 'undefined') {
        Microsoft.Office.WebExtension.useShortNamespace(false);
    }
    else {
        Office.useShortNamespace(false);
    }
    write('Office alias is now ' + typeof Office);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

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

## Historial de compatibilidad


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint Online.|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado compatibilidad para llamar a este método en los complementos de contenido para Access.|
|1.0|Agregado|
