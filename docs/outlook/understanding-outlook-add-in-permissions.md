
# Especificar permisos para el acceso de los complementos de Outlook al buzón del usuario

Los complementos de Outlook especifican el nivel de permisos requerido en su manifiesto. Los niveles disponibles son  **Restricted**,  **ReadItem**,  **ReadWriteItem** o **ReadWriteMailbox**. Estos niveles de permisos son acumulativos:  **Restringido** es el nivel más bajo y cada nivel superior incluye los permisos de todos los niveles que estén por debajo. **ReadWriteMailbox** incluye todos los permisos admitidos.

Puede ver los permisos que solicita un complemento de correo antes de instalarlo desde la Tienda Office. También puede ver los permisos necesarios de los complementos instalados en el Centro de administración de Exchange.


## Permiso restringido


El permiso  **Restringido** es el nivel más básico de los permisos. Especifique **Restricted** en el elemento [Permissions](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) del manifiesto para pedir este permiso. Si un complemento de correo no pide un permiso específico en su manifiesto, Outlook le asigna este permiso de forma predeterminada.


### Se puede


- [Obtener solo entidades específicas](../outlook/match-strings-in-an-item-as-well-known-entities.md) (número de teléfono, dirección, dirección URL) del asunto o el cuerpo del elemento.
    
- Especificar una [regla de activación ItemIs](../outlook/manifests/activation-rules.md#itemis-rule) que necesite que el elemento actual en un formulario de lectura o redacción sea de un determinado tipo de elemento, o bien una [regla ItemHasKnownEntity](../outlook/match-strings-in-an-item-as-well-known-entities.md) que coincida con cualquier subconjunto más pequeño de entidades conocidas (número de teléfono, dirección, dirección URL) del elemento seleccionado.
    
- Tener acceso a las propiedades y métodos que  **no** hacen referencia a información específica sobre el usuario o el elemento (vea la siguiente sección para ver una lista de los miembros que sí hacen referencia).
    

### No se puede


- Usar una regla [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) en la entidad de contacto, dirección de correo electrónico, sugerencia de reunión o sugerencia de tarea.
    
- Usar las reglas [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx) o [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx).
    
- Acceder a los miembros de la siguiente lista que hacen referencia a la información sobre el usuario o el elemento. Si trata de acceder a los miembros de esta lista, se devolverá  **null** y un mensaje de error indicará que Outlook solicita al complemento de correo un permiso elevado.
    
      - [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.attachments](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.bcc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.body](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.cc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.from](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatches](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.organizer](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.resources](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.sender](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.to](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.userProfile](../../reference/outlook/Office.context.mailbox.userProfile.md)
    
  - [Body](../../reference/outlook/Body.md) y todos sus miembros secundarios
    
  - [Location](../../reference/outlook/Location.md) y todos sus miembros secundarios
    
  - [Recipients](../../reference/outlook/Recipients.md) y todos sus miembros secundarios
    
  - [Subject](../../reference/outlook/Subject.md) y todos sus miembros secundarios
    
  - [Time](../../reference/outlook/Time.md) y todos sus miembros secundarios
    

## Permiso ReadItem


El permiso  **ReadItem** es el siguiente nivel de permiso en el modelo de permisos. Especifique **ReadItem** en el elemento **Permissions** del manifiesto para solicitar este permiso.


### Se puede


- [Leer todas las propiedades](../outlook/item-data.md) del elemento actual en un formulario de lectura o un [formulario de redacción](../outlook/get-and-set-item-data-in-a-compose-form.md) (por ejemplo, [item.to](../../reference/outlook/Office.context.mailbox.item.md) en un formulario de lectura e [item.to.getAsync](../../reference/outlook/Recipients.md) en uno de redacción).
    
- [Obtener un token de devolución de llamada para obtener los datos adjuntos del elemento](../outlook/get-attachments-of-an-outlook-item.md) o el elemento completo.
    
- [Escribir propiedades personalizadas](http://msdn.microsoft.com/library/30217d63-7615-4f3f-8618-c91e4e60cd43%28Office.15%29.aspx) definidas por el complemento en ese elemento.
    
- [Obtener todas las entidades conocidas existentes](../outlook/match-strings-in-an-item-as-well-known-entities.md), no solo un subconjunto, del asunto o el cuerpo del elemento.
    
- Usar todas las [entidades conocidas](../outlook/manifests/activation-rules.md#itemhasknownentity-rule) en [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) o las [expresiones regulares](../outlook/manifests/activation-rules.md#itemhasregularexpressionmatch-rule) en [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) En este ejemplo se sigue el esquema v1.1. Muestra una regla que activa el complemento si una o varias de las entidades conocidas están en el asunto o el cuerpo del mensaje seleccionado:
    

```XML
<Permissions>ReadItem</Permissions>
    <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="PhoneNumber" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="MeetingSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="TaskSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="EmailAddress" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
</Rule>
```


### No se puede

Obtener acceso a  **mailbox.makeEWSRequestAsync** o a cualquiera de los siguientes métodos de escritura:


- [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.bcc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.bcc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.body.prependAsync](../../reference/outlook/Body.md)
    
- [item.body.setAsync](../../reference/outlook/Body.md)
    
- [item.body.setSelectedDataAsync](../../reference/outlook/Body.md)
    
- [item.cc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.cc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.end.setAsync](../../reference/outlook/Time.md)
    
- [item.location.setAsync](../../reference/outlook/Location.md)
    
- [item.optionalAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.optionalAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.requiredAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.requiredAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.start.setAsync](../../reference/outlook/Time.md)
    
- [item.subject.setAsync](../../reference/outlook/Subject.md)
    
- [item.to.addAsync](../../reference/outlook/Recipients.md)
    
- [item.to.setAsync](../../reference/outlook/Recipients.md)
    

## Permiso ReadWriteItem


Especifique  **ReadWriteItem** en el elemento **Permissions** del manifiesto para solicitar este permiso. Los complementos de correo activadas en los formularios de redacción que usen métodos de escritura ( **Message.to.addAsync** o **Message.to.setAsync**) deben usar como mínimo este nivel de permiso.


### Se puede


- [Leer y escribir todas las propiedades de nivel de elemento](../outlook/item-data.md) del elemento que se va a ver o redactar en Outlook.
    
- [Agregar o quitar datos adjuntos](../outlook/add-and-remove-attachments-to-an-item-in-a-compose-form.md) de ese elemento.
    
- Usar todos los demás miembros de la API de JavaScript para Office válidos para los complementos de correo, excepto  **Mailbox.makeEWSRequestAsync**.
    

### No se puede

Usar  **Mailbox.makeEWSRequestAsync**.


## Permiso ReadWriteMailbox


El permiso  **ReadWriteMailbox** es el nivel de permiso más elevado. Especifique **ReadWriteMailbox** en el elemento **Permissions** del manifiesto para solicitar este permiso.

Aparte de lo que el permiso  **lReadWriteItem** admite, cuando se usa **Mailbox.makeEWSRequestAsync**, puede acceder a las operaciones de servicios Web Exchange (EWS) admitidas para hacer lo siguiente:


- Leer y escribir todas las propiedades de cualquier elemento en el buzón del usuario.
    
- Crear, leer y escribir en una carpeta o elemento en ese buzón.
    
- Enviar un elemento desde ese buzón.
    
A través de  **mailbox.makeEWSRequestAsync**, puede tener acceso a las siguientes operaciones de EWS:


- [CopyItem](http://msdn.microsoft.com/en-us/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)
    
- [CreateFolder](http://msdn.microsoft.com/en-us/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)
    
- [CreateItem](http://msdn.microsoft.com/en-us/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)
    
- [FindConversation](http://msdn.microsoft.com/en-us/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)
    
- [FindFolder](http://msdn.microsoft.com/en-us/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)
    
- [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)
    
- [GetConversationItems](http://msdn.microsoft.com/en-us/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)
    
- [GetFolder](http://msdn.microsoft.com/en-us/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)
    
- [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)
    
- [MarkAsJunk](http://msdn.microsoft.com/en-us/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)
    
- [MoveItem](http://msdn.microsoft.com/en-us/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)
    
- [SendItem](http://msdn.microsoft.com/en-us/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)
    
- [UpdateFolder](http://msdn.microsoft.com/en-us/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)
    
- [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)
    
Si trata de usar una acción no admitida, recibirá una respuesta de error.


## Recursos adicionales



- [Privacidad, permisos y seguridad para los complementos de Outlook](../outlook/../../docs/develop/privacy-and-security.md)
    
- [Coincidencia de cadenas en un elemento de Outlook como entidades conocidas](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
