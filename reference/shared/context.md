
# Objeto Context
Representa el entorno en tiempo de ejecución del complemento y proporciona acceso a los objetos clave de la API.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project y Word|
|**Modificado por última vez en**|1.1|

```
Office.context
```


## Miembros

|||
|:-----|:-----|
|Nombre|Descripción|
|[commerceAllowed](../../reference/shared/office.context.commerceallowed.md)|Obtiene información sobre si el complemento se está ejecutando en una plataforma que admite vínculos a sistemas de pago externos.|
|[contentLanguage](../../reference/shared/office.context.contentlanguage.md)|Obtiene la configuración regional (de idioma) para los datos tal como se encuentra almacenada en el documento o el elemento.|
|[displayLanguage](../../reference/shared/office.context.displaylanguage.md)|Obtiene la configuración regional (de idioma) para la UI de la aplicación host.|
|[documento](../../reference/shared/office.context.document.md)|Obtiene un objeto que representa el documento con el que está interactuando el complemento de panel de tareas o de contenido.|
|[buzón de correo](../../reference/shared/office.context.mailbox.md)|Obtiene el objeto **mailbox** que proporciona acceso a los miembros de la API específicos para los complementos de Outlook.|
|[officeTheme](../../reference/shared/office.context.officetheme.md)|Proporciona acceso a las propiedades de los colores del tema de Office.|
|[ui](../../reference/shared/officeui)|Proporciona objetos y métodos que puede usar para crear y manipular componentes de la interfaz de usuario, como cuadros de diálogo.|
|[roamingSettings](../../reference/shared/office.context.roamingsettings.md)|Obtiene un objeto que representa la configuración personalizada guardada del complemento.|
|[touchEnabled](../../reference/shared/office.context.touchenabled.md)|Obtiene información sobre si el complemento se está ejecutando en una aplicación host de Office con funcionalidad táctil.|

## Comentarios

El objeto **Context** proporciona acceso a los objetos clave de la API de JavaScript para Office.


## Detalles de compatibilidad



|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas y Outlook|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se agregaron las propiedades **commerceAllowed** y **touchEnabledAdded** (solo para Excel, PowerPoint y Word en Office para iPad).|
|1.1|Se ha agregado compatibilidad para los complementos para Excel y Word en Office para iPad.|
|1.1|Para [contentLanguage](../../reference/shared/office.context.contentlanguage.md), [displayLanguage](../../reference/shared/office.context.displaylanguage.md) y [document](../../reference/shared/office.context.document.md), se ha agregado compatibilidad para los complementos de contenido para Access.|
|1.0|Agregado|
