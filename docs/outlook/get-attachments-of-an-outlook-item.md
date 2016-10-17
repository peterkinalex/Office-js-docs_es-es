
# <a name="get-attachments-of-an-outlook-item-from-the-server"></a>Obtener datos adjuntos de un elemento de Outlook desde el servidor

Un complemento de Outlook no puede pasar los datos adjuntos de un elemento seleccionado directamente al servicio remoto que se ejecuta en el servidor. En su lugar, lo que puede hacer es usar la API de datos adjuntos para enviar información sobre los datos adjuntos al servicio remoto. Luego, el servicio puede ponerse en contacto directamente con el servidor de Exchange para recuperar los datos adjuntos.

Para enviar información de datos adjuntos al servicio remoto, se usa la siguiente función y las siguientes propiedades:


- Propiedad [Office.context.mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md): proporciona la dirección URL de los servicios Web Exchange (EWS) en el servidor Exchange donde se hospeda el buzón. El servicio usa esta dirección URL para llamar al método [ExchangeService.GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx) de la [API administrada de EWS](http://msdn.microsoft.com/library/c2267733-6f4f-49e5-9614-1e4a24c3af1a%28Office.15%29.aspx) o a la operación [GetAttachment](http://msdn.microsoft.com/en-us/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) de EWS.
    
- Propiedad [Office.context.mailbox.item.attachments](../../reference/outlook/Office.context.mailbox.item.md): obtiene una matriz de objetos [AttachmentDetails](../../reference/outlook/simple-types.md), una por cada adjunto al elemento.
    
- Función [Office.context.mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md): realiza una llamada asincrónica al servidor Exchange que hospeda el buzón para obtener un token de devolución de llamada que el servidor devuelve al servidor Exchange para autenticar una solicitud de datos adjuntos.
    

## <a name="using-the-attachments-api"></a>Usar la API de datos adjuntos


Para usar la API de datos adjuntos para obtener datos adjuntos de un buzón Exchange, siga estos pasos: 


1. Muestre el complemento cuando el usuario esté visualizando un mensaje o una cita que contenga datos adjuntos.
    
2. Obtenga el token de devolución de llamada del servidor Exchange.
    
3. Envíe el token de devolución de llamada y la información de datos adjuntos al servicio remoto.
    
4. Obtenga los datos adjuntos del servidor Exchange usando el método  **ExchangeService.GetAttachments** o la operación **GetAttachment**.
    
Estos pasos se explican en detalle en las siguientes secciones con el código del ejemplo de muestra [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments).


 >**Nota**  El código de estos ejemplos se ha abreviado para destacar la información de los datos adjuntos. El ejemplo contiene más código que sirve para autenticar el complemento en el servidor remoto y administrar el estado de la solicitud.


### <a name="activate-the-add-in"></a>Activar el complemento


Puede usar una regla [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx) en el archivo de manifiesto del complemento para mostrar su complemento de correo cuando el elemento seleccionado tenga datos adjuntos, como se muestra en el siguiente ejemplo.


```XML
<Rule xsi:type="ItemHasAttachment" />
```


### <a name="get-a-callback-token"></a>Obtener un token de devolución de llamada


El objeto [Office.context.mailbox](../../reference/outlook/Office.context.mailbox.md) proporciona la función **getCallbackTokenAsync** para obtener un token que el servidor remoto puede usar para la autenticación con el servidor Exchange. En el siguiente código se muestra una función en un complemento que inicia la solicitud asincrónica para obtener el token de devolución de llamada y la función de devolución de llamada que obtiene la respuesta. El token de devolución de llamada se almacena en el objeto de solicitud de servicio que se define en la siguiente sección.


```
function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "" {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
};
function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status === "succeeded") {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
};
```


### <a name="send-attachment-information-to-the-remote-service"></a>Enviar información de datos adjuntos al servicio remoto


El servicio remoto al que el complemento llama define los detalles sobre cómo se debe enviar la información de datos adjuntos al servicio. En este ejemplo, el servicio remoto es una aplicación API web creada mediante Visual Studio 2013. El servicio remoto espera que la información de datos adjuntos esté en un objeto JSON. Con el siguiente código se inicializa un objeto que contiene la información de datos adjuntos.


```
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
serviceRequest = new Object();
serviceRequest.attachmentToken = "";
serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
serviceRequest.attachments = new Array();
```

La propiedad  `Office.context.mailbox.item.attachments` contiene una colección de objetos **AttachmentDetails**, uno por cada dato adjunto del elemento. La mayoría de la veces, el complemento puede simplemente pasar la propiedad de identificador de datos adjuntos de un objeto  **AttachmentDetails** al servicio remoto. Si el servicio remoto necesita más detalles sobre los datos adjuntos, se puede pasar todo el objeto **AttachmentDetails** o parte de él. Con el siguiente código se define un método que coloca la matriz de **AttachmentDetails** entera en el objeto `serviceRequest` y se envía una solicitud al servicio remoto.




```js
    function makeServiceRequest() {
      // Format the attachment details for sending.
      for (var i = 0; i < mailbox.item.attachments.length; i++) {
        serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i].$0_0));
      }

      $.ajax({
        url: '../../api/Default',
        type: 'POST',
        data: JSON.stringify(serviceRequest),
        contentType: 'application/json;charset=utf-8'
      }).done(function (response) {
        if (!response.isError) {
          var names = "<h2>Attachments processed using " +
                        serviceRequest.service +
                        ": " +
                        response.attachmentsProcessed +
                        "</h2>";
          for (i = 0; i < response.attachmentNames.length; i++) {
            names += response.attachmentNames[i] + "<br />";
          }
          document.getElementById("names").innerHTML = names;
        } else {
          app.showNotification("Runtime error", response.message);
        }
      }).fail(function (status) {

      }).always(function () {
        $('.disable-while-sending').prop('disabled', false);
      })
    };

```


### <a name="get-the-attachments-from-the-exchange-server"></a>Obtener los datos adjuntos del servidor Exchange


El servicio remoto usa el método [GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx) de la API administrada de EWS o la operación [GetAttachment](http://msdn.microsoft.com/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) de EWS para recuperar los datos adjuntos del servidor. La aplicación de servicio necesita dos objetos para deserializar la cadena JSON y convertirla en objetos .NET Framework que se puedan usar en el servidor. En el código siguiente se muestran las definiciones de los objetos de deserialización.


```C#



namespace AttachmentsSample
{
  public class AttachmentSampleServiceRequest
  {
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public string service { get; set; }
    public AttachmentDetails [] attachments { get; set; }
  }

  public class AttachmentDetails
  {
    public string attachmentType { get; set; }
    public string contentType { get; set; }
    public string id { get; set; }
    public bool isInline { get; set; }
    public string name { get; set; }
    public int size { get; set; }
  }
}
```


#### <a name="use-the-ews-managed-api-to-get-the-attachments"></a>Uso de la API administrada de EWS para obtener datos adjuntos

Si utiliza la [API administrada de EWS](http://go.microsoft.com/fwlink/?LinkID=255472) en su servicio remoto, puede usar el método [GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx), el cual creará, enviará y recibirá una solicitud EWS SOAP para obtener los datos adjuntos. Se recomienda usar la API administrada de EWS porque requiere menos líneas de código y tiene una interfaz para realizar llamadas a EWS más intuitiva. El código siguiente realiza una solicitud para recuperar todos los datos adjuntos y devuelve el número y los nombres de los datos adjuntos procesados.


```C#
    private AttachmentSampleServiceResponse GetAtttachmentsFromExchangeServerUsingEWSManagedApi(AttachmentSampleServiceRequest request)
    {
      var attachmentsProcessedCount = 0;
      var attachmentNames = new List<string>();

      // Create an ExchangeService object, set the credentials and the EWS URL.
      ExchangeService service = new ExchangeService();
      service.Credentials = new OAuthCredentials(request.attachmentToken);
      service.Url = new Uri(request.ewsUrl);

      var attachmentIds = new List<string>();

      foreach (AttachmentDetails attachment in request.attachments)
      {
        attachmentIds.Add(attachment.id);
      }

      // Call the GetAttachments method to retrieve the attachments on the message.
      // This method results in a GetAttachments EWS SOAP request and response
      // from the Exchange server.
      var getAttachmentsResponse =
        service.GetAttachments(attachmentIds.ToArray(),
                               null,
                               new PropertySet(BasePropertySet.FirstClassProperties,
                                               ItemSchema.MimeContent));

      if (getAttachmentsResponse.OverallResult == ServiceResult.Success)
      {
        foreach (var attachmentResponse in getAttachmentsResponse)
        {
          attachmentNames.Add(attachmentResponse.Attachment.Name);

          // Write the content of each attachment to a stream.
          if (attachmentResponse.Attachment is FileAttachment)
          {
            FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
            Stream s = new MemoryStream(fileAttachment.Content);
            // Process the contents of the attachment here.
          }

          if (attachmentResponse.Attachment is ItemAttachment)
          {
            ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
            Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);
            // Process the contents of the attachment here.
          }

          attachmentsProcessedCount++;
        }
      }

      // Return the names and number of attachments processed for display
      // in the add-in UI.
      var response = new AttachmentSampleServiceResponse();
      response.attachmentNames = attachmentNames.ToArray();
      response.attachmentsProcessed = attachmentsProcessedCount;

      return response;
    }


```


#### <a name="use-ews-to-get-the-attachments"></a>Uso de EWS para obtener datos adjuntos

Si utiliza EWS en el servicio remoto, tendrá que crear una solicitud SOAP [GetAttachment](http://msdn.microsoft.com/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) para obtener los datos adjuntos del servidor Exchange. El código siguiente devuelve una cadena que proporciona la solicitud SOAP. El servicio remoto usa el método **String.Format** para insertar el Id. del dato adjunto en la cadena.


```C#
    private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";

```

Por último, el método siguiente usa una solicitud  **GetAttachment** de EWS para obtener los datos adjuntos del servidor Exchange. Esta implementación realiza una sola solicitud para cada dato adjunto y devuelve el número de datos adjuntos procesados. Cada respuesta se procesa en un método **ProcessXmlResponse** distinto, definido a continuación.




```C#
    private AttachmentSampleServiceResponse GetAttachmentsFromExchangeServerUsingEWS(AttachmentSampleServiceRequest request)
    {
      var attachmentsProcessedCount = 0;
      var attachmentNames = new List<string>();

      foreach (var attachment in request.attachments)
      {
        // Prepare a web request object.
        HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
        webRequest.Headers.Add("Authorization",
          string.Format("Bearer {0}", request.attachmentToken));
        webRequest.PreAuthenticate = true;
        webRequest.AllowAutoRedirect = false;
        webRequest.Method = "POST";
        webRequest.ContentType = "text/xml; charset=utf-8";

        // Construct the SOAP message for the GetAttachment operation.
        byte[] bodyBytes = Encoding.UTF8.GetBytes(
          string.Format(GetAttachmentSoapRequest, attachment.id));
        webRequest.ContentLength = bodyBytes.Length;

        Stream requestStream = webRequest.GetRequestStream();
        requestStream.Write(bodyBytes, 0, bodyBytes.Length);
        requestStream.Close();

        // Make the request to the Exchange server and get the response.
        HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

        // If the response is okay, create an XML document from the reponse
        // and process the request.
        if (webResponse.StatusCode == HttpStatusCode.OK)
        {
          var responseStream = webResponse.GetResponseStream();

          var responseEnvelope = XElement.Load(responseStream);

          // After creating a memory stream containing the contents of the 
          // attachment, this method writes the XML document to the trace output.
          // Your service would perform it's processing here.
          if (responseEnvelope != null)
          {
            var processResult = ProcessXmlResponse(responseEnvelope);
            attachmentNames.Add(string.Format("{0} {1}", attachment.name, processResult));

          }

          // Close the response stream.
          responseStream.Close();
          webResponse.Close();

        }
        // If the response is not OK, return an error message for the 
        // attachment.
        else
        {
          var errorString = string.Format("Attachment \"{0}\" could not be processed. " +
            "Error message: {1}.", attachment.name, webResponse.StatusDescription);
          attachmentNames.Add(errorString);
        }
        attachmentsProcessedCount++;
      }

      // Return the names and number of attachments processed for display
      // in the add-in UI.
      var response = new AttachmentSampleServiceResponse();
      response.attachmentNames = attachmentNames.ToArray();
      response.attachmentsProcessed = attachmentsProcessedCount;

      return response;
    }

```

Cada respuesta de la operación  **GetAttachment** se envía al método **ProcessXmlResponse**. Este método comprueba si hay errores en la respuesta. Si no encuentra ninguno, procesa los datos adjuntos del archivo y los datos adjuntos del elemento. El método  **ProcessXmlResponse** realiza el grueso del trabajo de procesamiento de los datos adjuntos.




```C#
    // This method processes the response from the Exchange server.
    // In your application the bulk of the processing occurs here.
    private string ProcessXmlResponse(XElement responseEnvelope)
    {
      // First, check the response for web service errors.
      var errorCodes = from errorCode in responseEnvelope.Descendants
                       ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                       select errorCode;
      // Return the first error code found.
      foreach (var errorCode in errorCodes)
      {
        if (errorCode.Value != "NoError")
        {
          return string.Format("Could not process result. Error: {0}", errorCode.Value);
        }
      }

      // No errors found, proceed with processing the content.
      // First, get and process file attachments.
      var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                        ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                            select fileAttachment;
      foreach(var fileAttachment in fileAttachments)
      {
        var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
        var fileData = System.Convert.FromBase64String(fileContent.Value);
        var s = new MemoryStream(fileData);
        // Process the file attachment here. 
      }

      // Second, get and process item attachments.
      var itemAttachments = from itemAttachment in responseEnvelope.Descendants
                            ("{http://schemas.microsoft.com/exchange/services/2006/types}ItemAttachment")
                            select itemAttachment;
      foreach(var itemAttachment in itemAttachments)
      {
        var message = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Message");
        if (message != null)
        {
         // Process a message here.
          break;
        }
        var calendarItem = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarItem");
        if (calendarItem != null)
        {
          // Process calendar item here.
          break;
        }
        var contact = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Contact");
        if (contact != null)
        {
          // Process contact here.
          break;
        }
        var task = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Tontact");
        if (task != null)
        {
          // Process task here.
          break;
        }
        var meetingMessage = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingMessage");
        if (meetingMessage != null)
        {
          // Process meeting message here.
          break;
        }
        var meetingRequest = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingRequest");
        if (meetingRequest != null)
        {
          // Process meeting request here.
          break;
        }
        var meetingResponse = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingResponse");
        if (meetingResponse != null)
        {
          // Process meeting response here.
          break;
        }
        var meetingCancellation = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingCancellation");
        if (meetingCancellation != null)
        {
          // Process meeting cancellation here.
          break;
        }
      }
     
      return string.Empty;
    }

```


## <a name="additional-resources"></a>Recursos adicionales



- [Crear complementos de Outlook para formularios de lectura](../outlook/read-scenario.md)
    
- 
  [Explorar la API administrada de EWS, EWS y servicios web de Exchange](http://msdn.microsoft.com/library/0bc6f81d-cc10-42b0-ba5d-6f22ff55d51c%28Office.15%29.aspx)
    
- 
  [Empezar a trabajar con aplicaciones cliente de la API administrada de EWS](http://msdn.microsoft.com/library/c2267733-6f4f-49e5-9614-1e4a24c3af1a%28Office.15%29.aspx)
    
- [Outlook-Power-Hour_Code-Samples](https://github.com/OfficeDev/Outlook-Power-Hour-Code-Samples): `MyAttachments` y `AttachmentsDemo`
    
