
# <a name="context.contentlanguage-property"></a>Propiedad Context.contentLanguage
 Obtiene la configuración regional (de idioma) especificada por el usuario para editar el documento o el elemento.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project y Word|
|**Modificado por última vez en**|1.1|

```
var myContentLang = Office.context.contentLanguage;
```


## <a name="return-value"></a>Valor devuelto

Una **string** con el formato de etiqueta de idioma RFC 1766, como `en-US`.


## <a name="remarks"></a>Observaciones

El valor **contentLanguage** refleja la configuración de **Idioma de edición** que se ha especificado desde **Archivo**  >  **Opciones**  >  **Idioma**, en la aplicación host de Office.

En el caso de los complementos de contenido para las aplicaciones web de Access, la propiedad **contentLanguage** obtiene la referencia cultural del complemento (por ejemplo, "es-ES").


## <a name="example"></a>Ejemplo




```js
function sayHelloWithContentLanguage() {
    var myContentLanguage = Office.context.contentLanguage;
    switch (myContentLanguage) {
        case 'en-US':
            write('Hello!');
            break;
        case 'en-NZ':
            write('G\'day mate!');
            break;
    }
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Project**|v|||
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para PowerPoint Online.|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Se ha agregado el acceso a esta API a los complementos de contenido para Access.|
|1.0|Agregado|
