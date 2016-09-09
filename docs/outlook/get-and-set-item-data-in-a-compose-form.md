
# Obtener y definir datos de elementos en un formulario de redacción de Outlook
Aprenda a obtener o establecer varias propiedades de un elemento en un complemento de Outlook de un escenario de redacción, incluidos los destinatarios, el asunto, el cuerpo, la hora y la ubicación de la cita.




## Obtener y configurar propiedades de elemento para un complemento de redacción


En un formulario de redacción, puede obtener la mayoría de las propiedades expuestas en el mismo tipo de elemento que en un formulario de lectura (como asistentes, destinatarios, asunto y cuerpo), y también algunas propiedades adicionales que solo son relevantes en un formulario de redacción, pero no en uno de lectura (cuerpo, CCO). 

Para la mayoría de estas propiedades, como es posible que un complemento de Outlook y el usuario modifiquen la misma propiedad en la interfaz de usuario al mismo tiempo, los métodos para obtenerlas y configurarlas son asincrónicos. En la tabla 1 se muestran las propiedades de nivel de elemento y los métodos asincrónicos correspondientes para obtenerlas y configurarlas en un formulario de redacción. Las propiedades [item.itemType](../../reference/outlook/Office.context.mailbox.item.md) e [item.conversationId](../../reference/outlook/Office.context.mailbox.item.md) son excepciones porque los usuarios no pueden modificarlas. Puede obtenerlas mediante programación de la misma forma en un formulario de redacción y en uno de lectura, directamente del objeto primario.

Además de obtener acceso a propiedades de elemento en la API de JavaScript para Office, también puede tener acceso a propiedades de nivel de elemento mediante los servicios Web Exchange (EWS). Con el permiso  **ReadWriteMailbox** puede usar el método [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) para obtener acceso a las operaciones de EWS, [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) y [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) para obtener y establecer más propiedades de uno o varios elementos en el buzón del usuario. **makeEwsRequestAsync** está disponible tanto en formularios de redacción como de lectura. Si quiere más información sobre el permiso **ReadWriteMailbox** y sobre el acceso a EWS a través de la plataforma de Complementos de Office, vea [Especificar permisos para el acceso de los complementos de Outlook al buzón del usuario](../outlook/understanding-outlook-add-in-permissions.md) y [Llamar a servicios web desde un complemento de Outlook](../outlook/web-services.md).


**Tabla 1. Métodos asincrónicos para obtener o establecer propiedades de elemento en un formulario de redacción**


|**Propiedad**|**Tipo de propiedad**|**Método asincrónico para obtener**|**Métodos asincrónicos para establecer**|
|:-----|:-----|:-----|:-----|
|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|[Destinatarios](../../reference/outlook/Recipients.md)|[Recipients.getAsync](../../reference/outlook/Recipients.md)|[Recipients.addAsync](../../reference/outlook/Recipients.md)[Recipients.setAsync](../../reference/outlook/Recipients.md)|
|[cuerpo](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body.getAsync](../../reference/outlook/Body.md)|[Body.prependAsync](../../reference/outlook/Body.md)[Body.setAsync](../../reference/outlook/Body.md)[Body.setSelectedDataAsync](../../reference/outlook/Body.md)|
|[cc](../../reference/outlook/Office.context.mailbox.item.md)|Destinatarios|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[fin](../../reference/outlook/Office.context.mailbox.item.md)|[Hora](../../reference/outlook/Time.md)|[Time.getAsync](../../reference/outlook/Time.md)|[Time.setAsync](../../reference/outlook/Time.md)|
|[location](../../reference/outlook/Office.context.mailbox.item.md)|[Ubicación](../../reference/outlook/Location.md)|[Location.getAsync](../../reference/outlook/Location.md)|[Location.setAsync](../../reference/outlook/Location.md)|
|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|Destinatarios|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|Destinatarios|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[empezar](../../reference/outlook/Office.context.mailbox.item.md)|Hora|Time.getAsync|Time.setAsync|
|[subject](../../reference/outlook/Office.context.mailbox.item.md)|[Tema](../../reference/outlook/Subject.md)|[Subject.getAsync](../../reference/outlook/Subject.md)|[Subject.setAsync](../../reference/outlook/Subject.md)|
|[a](../../reference/outlook/Office.context.mailbox.item.md)|Destinatarios|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|



## Recursos adicionales



- [Crear complementos de Outlook para formularios de redacción](../outlook/compose-scenario.md)
    
- [Comprender los permisos de los complementos de Outlook](../outlook/understanding-outlook-add-in-permissions.md)
    
- [Llamar a servicios web desde un complemento de Outlook](../outlook/web-services.md)
    
- [Obtención y definición de datos de elementos de Outlook en los formularios de lectura o redacción](../outlook/item-data.md)
    


