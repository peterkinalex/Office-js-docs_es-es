
# Obtención y definición de datos de elementos de Outlook en los formularios de lectura o redacción

A partir de la versión 1.1 del esquema de manifiestos de los Complementos de Office, Outlook puede activar complementos cuando el usuario ve o redacta un elemento. Las propiedades que están disponibles para un complemento en el elemento cambian según si el complemento se activa en un formulario de lectura o de redacción. Por ejemplo, las propiedades [dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) y [dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md) están definidas solo para un elemento que ya se envió (elemento que posteriormente se visualiza en un formulario de lectura), pero no cuando se está creando el elemento (en un formulario de redacción). Otro ejemplo es la propiedad [bcc](../../reference/outlook/Office.context.mailbox.item.md), que solo sirve cuando se está creando un mensaje (en un formulario de redacción) y el usuario no puede obtener acceso a él en un formulario de lectura.

En la tabla 1 se muestran las propiedades de nivel de elemento en la API de JavaScript para Office que están disponibles en cada uno de los modos de lectura y redacción de los complementos de correo. Por lo general, las propiedades disponibles en los formularios de lectura son de solo lectura y aquellas disponibles en los formularios de redacción son de lectura y escritura, con la excepción de las propiedades [itemId](../../reference/outlook/Office.context.mailbox.item.md) y [conversationId](../../reference/outlook/Office.context.mailbox.item.md), que siempre son de solo lectura. Para el resto de las propiedades de nivel de elemento disponibles en los formularios de redacción, como posiblemente el complemento y el usuario puedan leer o escribir la misma propiedad al mismo tiempo, los métodos para obtener o configurar estas propiedades en modo de redacción son asincrónicos y, por lo tanto, el tipo de objetos devueltos también es diferente en los formularios de redacción y en los de lectura. Para obtener más información sobre cómo usar los métodos asincrónicos para obtener o configurar propiedades de nivel de elemento en modo de redacción, vea [Obtener y definir datos de elementos en un formulario de redacción de Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md).


**Tabla 1. Propiedades de los elementos disponibles en los formularios de redacción y de lectura**


|**Tipo de elemento**|**Propiedad**|**Tipo de propiedad en formularios de lectura**|**Tipo de propiedad en formularios de redacción**|
|:-----|:-----|:-----|:-----|
|Citas y mensajes|[dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md)|Objeto  **Date** de JavaScript|Propiedad no disponible|
|Citas y mensajes|[dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|Objeto  **Date** de JavaScript|Propiedad no disponible|
|Citas y mensajes|[itemClass](../../reference/outlook/Office.context.mailbox.item.md)|String|Propiedad no disponible|
|Citas y mensajes|[itemId](../../reference/outlook/Office.context.mailbox.item.md)|String|Propiedad no disponible|
|Citas y mensajes|[itemType](../../reference/outlook/Office.context.mailbox.item.md)|Cadena en enumeración [ItemType](../../reference/outlook/Office.MailboxEnums.md)|Propiedad no disponible|
|Citas y mensajes|[datos adjuntos](../../reference/outlook/Office.context.mailbox.item.md)|[AttachmentDetails](../../reference/outlook/simple-types.md)|Propiedad no disponible|
|Citas y mensajes|[cuerpo](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body](../../reference/outlook/Body.md)|
|Citas|[fin](../../reference/outlook/Office.context.mailbox.item.md)|Objeto  **Date** de JavaScript|[Hora](../../reference/outlook/Time.md)|
|Citas|[location](../../reference/outlook/Office.context.mailbox.item.md)|String|[Ubicación](../../reference/outlook/Location.md)|
|Citas y mensajes|[normalizedSubject](../../reference/outlook/Office.context.mailbox.item.md)|String|Propiedad no disponible|
|Citas|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|[EmailAddressDetails](../../reference/outlook/simple-types.md)|[Destinatarios](../../reference/outlook/Recipients.md)|
|Citas|[organizer](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Propiedad no disponible|
|Citas|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Destinatarios|
|Citas|[recursos](../../reference/outlook/Office.context.mailbox.item.md)|String|Propiedad no disponible|
|Citas|[empezar](../../reference/outlook/Office.context.mailbox.item.md)|Objeto  **Date** de JavaScript|Hora|
|Citas y mensajes|[subject](../../reference/outlook/Office.context.mailbox.item.md)|String|[Tema](../../reference/outlook/Subject.md)|
|Mensajes|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|Propiedad no disponible|Destinatarios|
|Mensajes|[cc](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Destinatarios|
|Mensajes|[conversationId](../../reference/outlook/Office.context.mailbox.item.md)|String|Cadena (solo lectura)|
|Mensajes|[desde](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Propiedad no disponible|
|Mensajes|[internetMessageId](../../reference/outlook/Office.context.mailbox.item.md)|Entero|Propiedad no disponible|
|Mensajes|[sender](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Propiedad no disponible|
|Mensajes|[a](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Destinatarios|

## Uso de los tokens de devolución de llamadas de Exchange Server desde un complemento de lectura


Si el complemento de Outlook está activado en formularios de lectura, puede obtener un token de devolución de llamada de Exchange. Este token se puede usar en el código del lado servidor para obtener acceso al elemento completo a través de los servicios Web Exchange (EWS). Al especificar el permiso  **ReadItem** en el manifiesto del complemento, puede usar el método [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) para obtener el token de devolución de llamada de Exchange, la propiedad [mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md) para obtener la dirección URL del punto de conexión de EWS correspondiente al buzón del usuario e [item.itemId](../../reference/outlook/Office.context.mailbox.item.md) para obtener el identificador de EWS del elemento seleccionado. Luego puede pasar el token de devolución de llamada, la dirección URL del punto de conexión de EWS y el identificador del elemento de EWS al código del lado servidor para obtener acceso a la operación [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) y obtener más propiedades del elemento.


## Acceso a EWS desde un complemento de redacción o de lectura


También puede usar el método [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) para obtener acceso a las operaciones de los servicios Web Exchange (EWS)[GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) y [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) directamente desde el complemento. Puede usar estas operaciones para obtener y establecer muchas propiedades de un elemento especificado. Este método está disponible para los complementos de Outlook independientemente de si el complemento se activó en un formulario de lectura o de redacción si se especifica el permiso **ReadWriteMailbox** en el manifiesto del complemento. Si quiere obtener más información sobre el uso de **makeEwsRequestAsync** para obtener acceso a las operaciones de EWS, vea [Llamar a servicios web desde un complemento de Outlook](../outlook/web-services.md).


## Recursos adicionales



- [Complementos de Outlook](../outlook/outlook-add-ins.md)
    
- [Obtener y definir datos de elementos en un formulario de redacción de Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Llamar a servicios web desde un complemento de Outlook](../outlook/web-services.md)
    


