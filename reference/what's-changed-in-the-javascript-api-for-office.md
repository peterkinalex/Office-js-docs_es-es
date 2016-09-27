
# Novedades en la API de JavaScript para Office
Con el fin de ampliar la funcionalidad de sus Complementos de Office, la API de JavaScript para Office se actualiza periódicamente con objetos, métodos, propiedades, eventos y enumeraciones nuevos y actualizados. Siga los vínculos siguientes para ver los miembros de la API nuevos y actualizados.

Para desarrollar complementos con nuevos miembros de API, necesita [actualizar los archivos de la API de JavaScript para Office en el proyecto](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

Para ver todos los miembros de la API incluidos los que no han cambiado desde actualizaciones anteriores, consulte [API de JavaScript para Office](../reference/javascript-api-for-office.md).


## API nuevas y actualizadas

 **Objetos nuevos y actualizados**


|**Object**|**Descripción**|**Versión agregada o actualizada**|
|:-----|:-----|:-----|
|[Elemento](../reference/outlook/Office.context.mailbox.item.md)|Actualizaciones y adiciones a:<br><ul><li><p>Los métodos <a href="../reference/outlook/Office.context.mailbox.item.md#getSelectedDataAsync" target="_blank">getSelectedDataAsync</a> y <a href="../reference/outlook/Office.context.mailbox.item.md#setSelectedDataAsync" target="_blank">setSelectedDataAsync</a> para permitir obtener la selección del usuario y sobrescribirla en el asunto y el cuerpo de un mensaje o cita.</p></li><li><p>Los métodos <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyAllForm" target="_blank">displayReplyAllForm</a> y <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyForm" target="_blank">displayReplyForm</a> para admitir la adición de datos adjuntos en el formulario de respuesta de una cita.</p></li></ul>|Mailbox 1.2|
|[Elemento](../reference/outlook/Office.context.mailbox.item.md)|Se actualizó para incluir métodos y campos para la creación de complementos de Outlook en modo de redacción. |1.1|
|[Binding](../reference/shared/binding.md)|Se actualizó para resultar compatible con enlaces de tablas en complementos de contenido para Access.|1.1|
|[Bindings](../reference/shared/bindings.bindings.md)|Se actualizó para resultar compatible con enlaces de tablas en complementos de contenido para Access.|1.1|
|[Body](../reference/outlook/Body.md)|Se agregó para habilitar la creación y edición del cuerpo de un mensaje o una cita en complementos de Outlook en modo de redacción.|1.1|
|[Documento](../reference/shared/document.md)|Actualizaciones y ediciones para: <ul><li><p>Admitir las propiedades <a href="http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00(Office.15).aspx" target="_blank">mode</a>, <a href="http://msdn.microsoft.com/library/77ba7daf-419f-44b6-8747-7fd5618b7053(Office.15).aspx" target="_blank">settings</a> y <a href="http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc(Office.15).aspx" target="_blank">url</a> en complementos de contenido para Access.</p></li><li><p>Obtener el documento como PDF con el método <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">getFileAsync</a> en complementos para PowerPoint y Word.</p></li><li><p>Obtener las propiedades de archivo con el método <a href="http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">getFileProperties</a> en complementos para Excel, PowerPoint y Word.</p></li><li><p>Navegar a ubicaciones y objetos dentro del documento con el método <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByIdAsync</a> en complementos para Excel y PowerPoint.</p></li><li><p>Obtener el identificador, el título y el índice de las diapositivas seleccionadas con el método <a href="http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">getSelectedDataAsync</a> (cuando especifique la nueva enumeración <span class="keyword">Office.CoercionType.SlideRange</span><a href="http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b(Office.15).aspx" target="_blank">coercionType</a>) en complementos para PowerPoint.</p></li></ul>|1.1|
|[Ubicación](../reference/outlook/Location.md)|Se agregó para habilitar la configuración de ubicación de una cita en complementos de Outlook en modo de redacción.|1.1|
|[Office](../reference/shared/office.md)|Se actualizó el método select para admitir la obtención de enlaces en complementos de contenido para Access.|1.1|
|[Destinatarios](../reference/outlook/Recipients.md)|Se agregó para habilitar la obtención y configuración de destinatarios de un mensaje o una cita en modo de redacción.|1.1|
|[Configuración](../reference/shared/document.settings.md)|Se actualizó para admitir la creación de configuraciones personalizadas en complementos de contenido para Access.|1.1|
|[Tema](../reference/outlook/Subject.md)|Se agregó para habilitar la obtención o configuración del asunto de un mensaje o una cita en complementos de Outlook en modo de redacción.|1.1|
|[Hora](../reference/outlook/Time.md)|Se agregó para habilitar la obtención y configuración de la hora de inicio y finalización de una cita en complementos de Outlook en modo de redacción.|1.1|



**Enumeraciones nuevas y actualizadas**


|**Object**|**Descripción**|**Versión**|
|:-----|:-----|:-----|
|[ActiveView](../reference/shared/activeview-enumeration.md)|Especifica el estado de la vista activa del documento (por ejemplo, si el usuario puede editar o no el documento).Se agregó para que los complementos para PowerPoint puedan determinar si los usuarios están viendo la presentación ( **Presentación con diapositivas**) o modificando diapositivas. |1.1|
|[CoercionType](../reference/shared/coerciontype-enumeration.md)|Se actualizó con  **Office.CoercionType.SlideRange** para admitir la obtención del intervalo de diapositivas seleccionado con el método **getSelectedDataAsync** en complementos para PowerPoint.|1.1|
|[EventType](../reference/shared/eventtype-enumeration.md)|Se actualizó para incluir el nuevo evento ActiveViewChanged.|1.1|
|[FileType](../reference/shared/filetype-enumeration.md)|Se actualizó para especificar el resultado en formato PDF.|1.1|
|[GoToType](../reference/shared/gototype-enumeration.md)|Se agregó para especificar el lugar u objeto del documento al que se debe ir.|1.1|

## Recursos adicionales


- [Referencias de esquema y API de complementos de Office](../reference/reference.md)
    
- [Office Add-ins](../docs/overview/office-add-ins.md)
    
