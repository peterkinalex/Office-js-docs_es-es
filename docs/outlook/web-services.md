
# <a name="call-web-services-from-an-outlook-add-in"></a>Llamar a servicios web desde un complemento de Outlook

Su complemento puede usar los servicios Web Exchange (EWS) desde un equipo que ejecute Exchange Server 2013, un servicio web disponible en el servidor que proporcione la ubicación de origen de la interfaz de usuario del complemento o un servicio web que esté disponible en Internet. En este artículo se proporciona un ejemplo que muestra cómo un complemento de Outlook puede solicitar información de EWS.

La forma de llamar a un servicio web varía en función del lugar en el que esté ubicado el servicio web. En la tabla 1 se enumeran las distintas formas que puede usar para llamar a un servicio web basado en la ubicación.


**Tabla 1: Formas de llamar a los servicios web desde un complemento de Outlook**


|**Ubicación del servicio web**|**Forma de llamar al servicio web**|
|:-----|:-----|
|El servidor Exchange que hospeda el buzón de correo del cliente|Use el método [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) para llamar a las operaciones de EWS que los complementos admiten. El servidor Exchange que hospeda el buzón de correo también expone EWS.|
|Servidor web que proporciona la ubicación de origen para la interfaz de usuario del complemento|Llame al servicio web con técnicas de JavaScript estándar. El código de JavaScript en el marco de la interfaz de usuario se ejecuta en el contexto del servidor web que proporciona la interfaz de usuario. Por lo tanto, puede llamar a los servicios web de ese servidor sin causar ningún error de scripting entre sitios.|
|Todas las demás ubicaciones|Cree un proxy para el servicio web en el servidor web que proporciona la ubicación de origen para la interfaz de usuario. Si no proporciona ningún proxy, los errores de scripting entre sitios impedirán la ejecución del complemento. Una manera de proporcionar un proxy es con JSON/P. Para obtener más información, consulte [Privacidad y seguridad de complementos para Office](../../docs/develop/privacy-and-security.md).|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a>Uso del método makeEwsRequestAsync para tener acceso a las operaciones de EWS


Puede usar el método [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) para crear una solicitud de EWS al servidor Exchange que hospeda el buzón del usuario.

Los EWS son compatibles con diversas operaciones que se llevan a cabo en un servidor Exchange, por ejemplo, operaciones del nivel de elemento para copiar, buscar, actualizar o enviar un elemento, y operaciones del nivel de carpeta para crear, obtener o actualizar una carpeta. Para realizar una operación de EWS, cree una solicitud SOAP XML para esa operación. Cuando finalice la operación, obtendrá una respuesta SOAP XML que contendrá los datos correspondientes a la operación. Las solicitudes y las respuestas SOAP de EWS siguen el esquema definido en el archivo Messages.xsd. Al igual que otros archivos de esquema de EWS, el archivo Message.xsd se encuentra en el directorio virtual de IIS que hospeda los EWS. 

Para usar el método  **makeEwsRequestAsync** para iniciar una operación de EWS, proporcione lo siguiente:


- El código XML para la solicitud SOAP para dicha operación de EWS, como un argumento para el parámetro  _data_
    
- Un método de devolución de llamada (como el argumento de  _callback_)
    
- Cualquier dato de entrada opcional para el método de devolución de llamada (como el argumento de  _userContext_)
    
Cuando la solicitud SOAP de EWS se complete, Outlook llamará al método de devolución de llamada con un argumento, que es un objeto [AsyncResult](../../reference/outlook/simple-types.md). El método de devolución de llamada puede obtener acceso a dos propiedades del objeto  **AsyncResult**: la propiedad  **value**, que contiene la respuesta SOAP XML de la operación de EWS y, de manera opcional, la propiedad  **asyncContext**, que contiene todos los datos que se pasan como parámetro  **userContext**. Por lo general, el método de devolución de llamada suele analizar el XML de la respuesta SOAP para obtener toda la información correspondiente y procesarla como corresponda.


## <a name="tips-for-parsing-ews-responses"></a>Sugerencias para analizar las respuestas de EWS


Al analizar la respuesta SOAP en una operación de EWS, tenga en cuenta los siguientes problemas según el explorador:


- Especifique el prefijo para un nombre de etiqueta cuando utilice el método DOM  **getElementsByTagName**, para incluir compatibilidad con Internet Explorer.
    
     **getElementsByTagName** se comporta de forma diferente según el tipo de explorador. Por ejemplo, una respuesta EWS puede contener el siguiente código XML (con formato y abreviado para fines de presentación):
    
```XML
      <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
    PropertyName="MyProperty" 
    PropertyType="String"/>
    <t:Value>{
    ...
    }</t:Value></t:ExtendedProperty>
```

 El siguiente código funcionaría en un explorador como Chrome para obtener el código XML delimitado por las etiquetas **ExtendedProperty**:

```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("ExtendedProperty");
```


   
 En Internet Explorer, tiene que incluir el prefijo `t:` del nombre de etiqueta, como se muestra a continuación:

```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("t:ExtendedProperty");
```

- Utilice la propiedad DOM  **textContent** para obtener el contenido de una etiqueta en una respuesta de EWS, tal como se muestra a continuación:
    
```
      content = $.parseJSON(value.textContent);
```

 Otras propiedades, como **innerHTML**, podrían no funcionar en Internet Explorer para algunas etiquetas en una respuesta de EWS.
    

## <a name="example"></a>Ejemplo


En el ejemplo siguiente se llama a  **makeEwsRequestAsync** para usar la operación [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) para obtener el asunto de un elemento. Este ejemplo incluye las tres funciones siguientes:


-  `getSubjectRequest`: usa el id. de un elemento como entrada y devuelve el XML para que la solicitud SOAP llame a **GetItem** para obtener el elemento especificado.
    
-  `sendRequest`: llama a `getSubjectRequest` para obtener la solicitud SOAP del elemento seleccionado y, después, pasa la solicitud SOAP y el método de devolución de llamada, `callback`, a **makeEwsRequestAsync** para obtener el asunto del elemento especificado.
    
-  `callback`: procesa la respuesta SOAP, que incluye la información y el asunto del elemento especificado.
    

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
'<?xml version="1.0" encoding="utf-8"?>' +
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
'  <soap:Header>' +
'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
'  </soap:Header>' +
'  <soap:Body>' +
'    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
'      <ItemShape>' +
'        <t:BaseShape>IdOnly</t:BaseShape>' +
'        <t:AdditionalProperties>' +
'            <t:FieldURI FieldURI="item:Subject"/>' +
'        </t:AdditionalProperties>' +
'      </ItemShape>' +
'      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
'    </GetItem>' +
'  </soap:Body>' +
'</soap:Envelope>';

   return result;
}





function sendRequest() {
   // Create a local variable that contains the mailbox.
   var mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
}


```


## <a name="ews-operations-that-add-ins-support"></a>Operaciones de EWS compatibles con los complementos


Los complementos de Outlook pueden tener acceso a un subconjunto de operaciones disponibles en EWS a través del método  **makeEwsRequestAsync**. Si no está familiarizado con las operaciones de EWS ni con el uso del método  **makeEwsRequestAsync** para obtener acceso a una operación, empiece con una solicitud SOAP de ejemplo para personalizar el argumento _data_. A continuación se describe el uso del método  **makeEwsRequestAsync**:


1. En el XML, sustituya todos los identificadores de elementos y los atributos de operaciones de EWS correspondientes por los valores adecuados.
    
2. Incluya la solicitud SOAP como un argumento del parámetro  _data_ de **makeEwsRequestAsync**.
    
3. Especifique un método de devolución de llamada y llame a  **makeEwsRequestAsync**.
    
4. En el método de devolución de llamada, compruebe los resultados de la operación en la respuesta SOAP.
    
5. Use los resultados de la operación de EWS de acuerdo con sus necesidades.
    
En la siguiente tabla se enumeran las operaciones de EWS compatibles con los complementos. Para ver ejemplos de solicitudes y respuestas de SOAP, elija el vínculo para cada operación. Para más información sobre operaciones de EWS, vea [Operaciones de EWS en Exchange](http://msdn.microsoft.com/library/cf6fd871-9a65-4f34-8557-c8c71dd7ce09%28Office.15%29.aspx).


**Tabla 2: Operaciones de EWS compatibles**


|**Operación de EWS**|**Descripción**|
|:-----|:-----|
|
  [Operación CopyItem](http://msdn.microsoft.com/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)|Copia los elementos especificados y coloca los elementos nuevos en una carpeta designada en el almacén de Exchange.|
|
  [Operación CreateFolder](http://msdn.microsoft.com/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)|Crea carpetas en la ubicación especificada en el almacén de Exchange.|
|
  [Operación CreateItem](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)|Crea los elementos especificados en el almacén de Exchange.|
|
  [Operación de FindConversation](http://msdn.microsoft.com/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)|Enumera una lista de conversaciones en la carpeta especificada en el almacén de Exchange.|
|
  [Operación FindFolder](http://msdn.microsoft.com/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)|Busca subcarpetas de una carpeta identificada y devuelve un conjunto de propiedades que describen el conjunto de subcarpetas.|
|
  [Operación FindItem](http://msdn.microsoft.com/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)|Identifica a los elementos que se encuentran en una carpeta especificada en el almacén de Exchange.|
|
  [Operación GetConversationItems](http://msdn.microsoft.com/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)|Obtiene uno o varios conjuntos de elementos que están organizados en nodos en una conversación.|
|
  [Operación GetFolder](http://msdn.microsoft.com/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)|Obtiene las propiedades y los contenidos especificados de las carpetas del almacén de Exchange.|
|
  [Operación GetItem](http://msdn.microsoft.com/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)|Obtiene las propiedades y los contenidos especificados de los elementos del almacén de Exchange.|
|
  [Operación MarkAsJunk](http://msdn.microsoft.com/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)|Mueve mensajes de correo a la carpeta Correo electrónico no deseado y agrega o quita de la lista de remitentes bloqueados a los remitentes de los mensajes, según corresponda.|
|
  [Operación MoveItem](http://msdn.microsoft.com/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)|Mueve elementos a una sola carpeta de destino en el almacén de Exchange.|
|
  [Operación SendItem](http://msdn.microsoft.com/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)|Envía mensajes de correo electrónico que se encuentran en el almacén de Exchange.|
|
  [Operación UpdateFolder](http://msdn.microsoft.com/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)|Modifica las propiedades de las carpetas existentes en el almacén de Exchange.|
|
  [Operación UpdateItem](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)|Modifica las propiedades de los elementos existentes en el almacén de Exchange.|

## <a name="authentication-and-permission-considerations-for-the-makeewsrequestasync-method"></a>Consideraciones de autenticación y permisos para el método makeEwsRequestAsync


Si usa el método  **makeEwsRequestAsync**, la solicitud se autentica con las credenciales de la cuenta de correo del usuario actual. El método  **makeEwsRequestAsync** administra las credenciales por usted, para que no tenga que proporcionarlas con la solicitud.


 >
  **Nota**  El administrador del servidor debe usar el cmdlet [New-WebServicesVirtualDirctory](http://technet.microsoft.com/en-us/library/bb125176.aspx) o [Set-WebServicesVirtualDirecory](http://technet.microsoft.com/en-us/library/aa997233.aspx) para establecer el parámetro _OAuthAuthentication_ como **true** en el directorio EWS del servidor de acceso de cliente para permitir que el método **makeEwsRequestAsync** realice la solicitud EWS.

El complemento tiene que especificar el permiso **ReadWriteMailbox** en su manifiesto de complemento para usar el método **makeEwsRequestAsync**. Para más información sobre cómo usar el permiso **ReadWriteMailbox**, vea la sección [Permiso ReadWriteMailbox](../outlook/understanding-outlook-add-in-permissions.md#readwritemailbox-permission) en [Información sobre los permisos del complemento de Outlook](../outlook/understanding-outlook-add-in-permissions.md).


## <a name="additional-resources"></a>Recursos adicionales



- [Complementos de Outlook](../outlook/outlook-add-ins.md)
    
- [Privacidad y seguridad de complementos para Office](../../docs/develop/privacy-and-security.md)
    
- [Abordar las limitaciones de la directiva de mismo origen en complementos para Office](../../docs/develop/addressing-same-origin-policy-limitations.md)
    
- 
  [Referencia EWS para Exchange](http://msdn.microsoft.com/library/2a873474-1bb2-4cb1-a556-40e8c4159f4a%28Office.15%29.aspx)
    
- 
  [Aplicaciones de correo para Outlook y EWS en Exchange](http://msdn.microsoft.com/library/821c8eb9-bb58-42e8-9a3a-61ca635cba59%28Office.15%29.aspx)
    
Si desea crear servicios back-end para complementos con ASP.NET Web API, consulte los recursos siguientes:


- [Crear un servicio web para un complemento para Office con ASP.NET Web API](http://blogs.msdn.com/b/officeapps/archive/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api.aspx)
    
- [Conceptos básicos sobre cómo compilar un servicio HTTP con ASP.NET Web API](http://www.asp.net/web-api)
    
