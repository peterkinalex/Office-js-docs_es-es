
# <a name="extract-entity-strings-from-an-outlook-item"></a>Extraer cadenas de entidad de un elemento de Outlook

En este artículo se describe cómo crear un complemento de Outlook de  **entidades de presentación** que extrae ejemplos de cadenas de entidades conocidas compatibles en el asunto y el cuerpo del elemento de Outlook seleccionado. Dicho elemento puede ser una cita, un mensaje de correo o una convocatoria, respuesta o cancelación de reunión. Entre las entidades compatibles se incluyen las siguientes:

- Dirección: dirección postal del país con al menos un subconjunto de los elementos: número de la calle, nombre de la calle, ciudad, estado y código postal.
    
- Contacto: información de contacto de una persona, en el contexto de otras entidades como una dirección o nombre empresarial.
    
- Dirección de correo electrónico: dirección de correo electrónico SMTP.
    
- Sugerencia de reunión: sugerencia de reunión, como una referencia a un evento. Tenga en cuenta que solo los mensajes admiten la extracción de sugerencias de reunión, no las citas.
    
- Número de teléfono: número de teléfono del país.
    
- Sugerencia de tarea: sugerencia de tarea, normalmente expresada en una frase que requiere una acción.
    
- Dirección URL.
    
La mayoría de estas entidades dependen del reconocimiento del lenguaje natural, que se basa en el aprendizaje automático de grandes cantidades de datos. Este reconocimiento no es determinista y, a veces, depende del contexto en el elemento de Outlook. Outlook activa el complemento de entidades cuando el usuario selecciona una cita, un mensaje de correo electrónico, una convocatoria de reunión, una respuesta o una cancelación para su visualización. Durante la inicialización, el complemento de entidades de ejemplo lee todas las instancias de las entidades compatibles del elemento actual. 

El complemento proporciona botones para que el usuario pueda elegir un tipo de entidad. Cuando el usuario selecciona una entidad, en el complemento se muestran instancias de la entidad seleccionada en el panel de complemento. En las secciones siguientes se proporciona una lista de los archivos de manifiesto XML, HTML y JavaScript del complemento de entidades, y se resalta el código que es compatible con la extracción de entidades correspondiente.

## <a name="xml-manifest"></a>Manifiesto XML


El complemento de entidades tiene dos reglas de activación unidas por una operación OR lógica. 


```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

Estas reglas indican que Outlook debe activar el complemento cuando el elemento seleccionado actualmente en el panel o inspector de lectura sea una cita o un mensaje (incluidos los mensajes de correo o las convocatorias, respuestas o cancelaciones de reunión).

El siguiente es el manifiesto del complementos de entidades. Usa la versión 1.1 del esquema para manifiestos de Complementos de Office.




```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
xsi:type="MailApp">
  <Id>6880A140-1C4F-11E1-BDDB-0800200C9A68</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>EN-US</DefaultLocale>
  <DisplayName DefaultValue="Display entities"/>
  <Description DefaultValue=
     "Display known entities on the selected item."/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify the web -->
        <!-- server where the HTML file is hosted. -->
        <SourceLocation DefaultValue=
          "http://webserver/default_entities/default_entities.html"/>
        <RequestedHeight>350</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <!-- Activate the add-in if the current item in Outlook is -->
  <!-- an email or appointment item. -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>
```


## <a name="html-implementation"></a>Implementación de HTML


El archivo HTML del complemento de entidades define los botones necesarios para que el usuario seleccione los tipos de entidad y otro botón para desactivar las instancias mostradas de una entidad. Incluye un archivo JavaScript, default_entities.js, que se describe más adelante en la sección [Implementación de JavaScript](#javascript-implementation). El archivo JavaScript incluye los controladores de eventos para cada uno de los botones.

Tenga en cuenta que todos los complementos de Outlook deben incluir office.js. El archivo HTML siguiente incluye la versión 1.1 de office.js en la red CDN. 

```html
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" >
    <title>standard_item_properties</title>
    <link rel="stylesheet" type="text/css" media="all" href="default_entities.css" />
    <script type="text/javascript" src="MicrosoftAjax.js"></script>
    <!-- Use the CDN reference to Office.js. -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    <script type="text/javascript"  src="default_entities.js"></script>
</head>

<body>
    <div id="container">
        <div id="button">
        <input type="button" value="clear" 
            onclick="myClearEntitiesBox();">
        <input type="button" value="Get Addresses" 
            onclick="myGetAddresses();">
        <input type="button" value="Get Contact Information" 
            onclick="myGetContacts();">
        <input type="button" value="Get Email Addresses" 
            onclick="myGetEmailAddresses();">
        <input type="button" value="Get Meeting Suggestions" 
            onclick="myGetMeetingSuggestions();">
        <input type="button" value="Get Phone Numbers" 
            onclick="myGetPhoneNumbers();">
        <input type="button" value="Get Task Suggestions" 
            onclick="myGetTaskSuggestions();">
        <input type="button" value="Get URLs" 
            onclick="myGetUrls();">
        </div>
        <div id="entities_box"></div>
    </div>
</body>
</html>
```


## <a name="style-sheet"></a>Hoja de estilos


El complemento de entidades usa un archivo CSS opcional, default_entities.css, para especificar el diseño de los resultados. A continuación, se muestra una lista del archivo CSS.


```css
*
{
    color: #FFFFFF;
    margin: 0px;
    padding: 0px;
    font-family: Arial, Sans-serif;
}
html 
{
    scrollbar-base-color: #FFFFFF;
    scrollbar-arrow-color: #ABABAB; 
    scrollbar-lightshadow-color: #ABABAB; 
    scrollbar-highlight-color: #ABABAB; 
    scrollbar-darkshadow-color: #FFFFFF; 
    scrollbar-track-color: #FFFFFF;
}
body
{
    background: #4E9258;
}
input
{
    color: #000000;
    padding: 5px;
}
span
{
    color: #FFFF00;
}
div#container
{
    height: 100%;
    padding: 2px;
    overflow: auto;
}
div#container td
{
    border-bottom: 1px solid #CCCCCC;
}
td.property-name
{
    padding: 0px 5px 0px 0px;
    border-right: 1px solid #CCCCCC;
}
div#meeting_suggestions
{
    border-top: 1px solid #CCCCCC;
}
```


## <a name="javascript-implementation"></a>Implementación de JavaScript


Las secciones restantes describen cómo esta muestra (archivo default_entities.js) extrae las entidades conocidas del asunto y el cuerpo del mensaje o de la cita que el usuario está viendo. 


## <a name="extracting-entities-upon-initialization"></a>Extracción de entidades en la inicialización


Tras el evento [Office.initialize](../../reference/shared/office.initialize.md), el complemento de entidades llama al método [getEntities](../../reference/outlook/Office.context.mailbox.item.md) del elemento actual. El método **getEntities** devuelve la variable global `_MyEntities`, una matriz de instancias de entidades compatibles. A continuación, se muestra el código JavaScript correspondiente.


```js
// Global variables
var _Item;
var _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}

```


## <a name="extracting-addresses"></a>Extracción de direcciones


Cuando el usuario hace clic en el botón **Obtener direcciones**, el controlador de eventos `myGetAddresses` obtiene una matriz de direcciones de la propiedad [addresses](../../reference/outlook/simple-types.md) del objeto `_MyEntities` (si se extrajo alguna dirección). Cada dirección extraída se almacena como una cadena en la matriz. `myGetAddresses` forma una cadena HTML local en .mdText para mostrar la lista de direcciones extraídas. A continuación, se muestra el código JavaScript correspondiente.


```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-contact-information"></a>Extracción de información de contacto


Cuando el usuario hace clic en el botón  **Obtener información de contacto**, el controlador de eventos  `myGetContacts` obtiene una matriz de contactos con la información de la propiedad [contacts](../../reference/outlook/simple-types.md) del objeto `_MyEntities` (si se extrajo alguno). Cada contacto extraído se almacena como un objeto [Contact](../../reference/outlook/simple-types.md) en la matriz. `myGetContacts` obtiene más información sobre cada contacto. Observe que el contexto determina si Outlook puede extraer un contacto de un elemento. Para ello, debería haber una firma al final de un mensaje de correo o alguno de los siguientes datos cerca del contacto:


- La cadena que representa el nombre del contacto de la propiedad [Contact.personName](../../reference/outlook/simple-types.md).
    
- La cadena que representa el nombre de la compañía asociado al contacto de la propiedad [Contact.businessName](../../reference/outlook/simple-types.md).
    
- La matriz de números de teléfono asociados con el contacto de la propiedad [Contact.phoneNumbers](../../reference/outlook/simple-types.md). Cada número de teléfono está representado en un objeto [PhoneNumber](../../reference/outlook/simple-types.md).
    
- Para cada miembro  **PhoneNumber** de la matriz de números de teléfono, la cadena que representa el número de teléfono de la propiedad [PhoneNumber.phoneString](../../reference/outlook/simple-types.md).
    
- La matriz de direcciones URL asociadas al contacto de la propiedad [Contact.urls](../../reference/outlook/simple-types.md). Cada dirección URL se representa como una cadena en un miembro de la matriz.
    
- La matriz de direcciones de correo electrónico asociadas con el contacto de la propiedad [Contact.emailAddresses](../../reference/outlook/simple-types.md). Cada dirección de correo está representada como una cadena en un miembro de la matriz.
    
- La matriz de direcciones postales asociadas con el contacto de la propiedad [Contact.addresses](../../reference/outlook/simple-types.md). Cada dirección postal está representada como una cadena en un miembro de la matriz.
    
 `myGetContacts` forma una cadena HTML local en `htmlText` para mostrar los datos de cada contacto. A continuación, se muestra el código JavaScript relacionado.




```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    var htmlText = "";

    // Gets an array of contacts and their information.
    var contactsArray = _MyEntities.contacts;
    for (var i = 0; i < contactsArray.length; i++)
    {
        // Gets the name of the person. The name is a string.
        htmlText += "Name : <span>" + contactsArray[i].personName +
            "</span><br/>";

        // Gets the company name associated with the contact.
        htmlText += "Business : <span>" + 
        contactsArray[i].businessName + "</span><br/>";

        // Gets an array of phone numbers associated with the 
        // contact. Each phone number is represented by a 
        // PhoneNumber object.
        var phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (var j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        var urlsArray = contactsArray[i].urls;
        for (var j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        var emailAddressesArray = contactsArray[i].emailAddresses;
        for (var j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        var addressesArray = contactsArray[i].addresses;
        for (var j = 0; j < addressesArray.length; j++)
        {
          htmlText += "Address : <span>" + addressesArray[j] + 
              "</span><br/>";
        }

        htmlText += "<hr/>";
        }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-email-addresses"></a>Extracción de direcciones de correo electrónico


Cuando el usuario hace clic en el botón  **Obtener direcciones de correo electrónico**, el controlador de eventos  `myGetEmailAddresses` obtiene una matriz de direcciones de correo electrónico SMTP de la propiedad [emailAddresses](../../reference/outlook/simple-types.md) del objeto `_MyEntities` (si se extrajo alguna). Cada dirección de correo electrónico extraída se almacena en una cadena de la matriz. `myGetEmailAddresses` forma una cadena HTML local en `htmlText` para mostrar la lista de direcciones de correo electrónico extraídas. A continuación, se muestra el código JavaScript correspondiente.


```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses() {
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++) {
        htmlText += "E-mail Address : <span>" + emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-meeting-suggestions"></a>Extracción de sugerencias de reunión


Cuando el usuario hace clic en el botón  **Obtener sugerencias de reunión**, el controlador de eventos  `myGetMeetingSuggestions` obtiene una matriz de sugerencias de reunión de la propiedad [meetingSuggestions](../../reference/outlook/simple-types.md) del objeto `_MyEntities` (si se extrajo alguna).


 >**Nota**  Solo los mensajes, y no las citas, admiten el tipo de entidad  **MeetingSuggestion**.

Cada sugerencia de reunión extraída se almacena en un objeto [MeetingSuggestion](../../reference/outlook/simple-types.md) dentro de la matriz. `myGetMeetingSuggestions` obtiene más información sobre cada sugerencia de reunión:


- La cadena que se identificó como una sugerencia de reunión de la propiedad [MeetingSuggestion.meetingString](../../reference/outlook/simple-types.md).
    
- La matriz de los asistentes a la reunión de la propiedad [MeetingSuggestion.attendees](../../reference/outlook/simple-types.md). Cada asistente está representado en un objeto [EmailUser](../../reference/outlook/simple-types.md).
    
- Por cada asistente, el nombre de la propiedad [EmailUser.displayName](../../reference/outlook/simple-types.md).
    
- Para cada asistente, la dirección SMTP de la propiedad [EmailUser.emailAddress](../../reference/outlook/simple-types.md).
    
- La cadena que representa la ubicación de la sugerencia de reunión de la propiedad [MeetingSuggestion.location](../../reference/outlook/simple-types.md).
    
- La cadena que representa el asunto de la sugerencia de reunión de la propiedad [MeetingSuggestion.subject](../../reference/outlook/simple-types.md).
    
- La cadena que representa la hora de inicio de la sugerencia de reunión de la propiedad [MeetingSuggestion.start](../../reference/outlook/simple-types.md).
    
- La cadena que representa la hora de finalización de la sugerencia de reunión de la propiedad [MeetingSuggestion.end](../../reference/outlook/simple-types.md).
    
 `myGetMeetingSuggestions` forma una cadena HTML local en `htmlText` para mostrar los datos de cada sugerencia de reunión. A continuación, se muestra el código JavaScript relacionado.




```js
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions() {
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++) {
        // Gets the string that was identified as a meeting suggestion.
        htmlText += "MeetingString : <span>" + meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++) {
            htmlText += "Attendee : ( ";

            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-phone-numbers"></a>Extracción de números de teléfono


Cuando el usuario hace clic en el botón  **Obtener números de teléfono**, el controlador de eventos  `myGetPhoneNumbers` obtiene una matriz de números de teléfono de la propiedad [phoneNumbers](../../reference/outlook/simple-types.md) del objeto `_MyEntities` (si se extrajo alguno). Cada número de teléfono extraído se almacena como un objeto [PhoneNumber](../../reference/outlook/simple-types.md) dentro de la matriz. `myGetPhoneNumbers` obtiene más información sobre cada número de teléfono:


- La cadena que representa el tipo de número de teléfono (por ejemplo, el teléfono particular) de la propiedad [PhoneNumber.type](../../reference/outlook/simple-types.md).
    
- La cadena que representa el número de teléfono actual de la propiedad [PhoneNumber.phoneString](../../reference/outlook/simple-types.md).
    
- La cadena que se identificó originalmente como el número de teléfono de la propiedad [PhoneNumber.originalPhoneString](../../reference/outlook/simple-types.md).
    
 `myGetPhoneNumbers` forma una cadena HTML local en `htmlText` para mostrar los datos de cada número de teléfono. A continuación, se muestra el código JavaScript relacionado.




```js
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
           "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
            phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="extracting-task-suggestions"></a>Extracción de sugerencias de tareas


Cuando el usuario hace clic en el botón  **Obtener sugerencias de tareas**, el controlador de eventos  `myGetTaskSuggestions` obtiene una matriz de sugerencias de tarea de la propiedad [taskSuggestions](../../reference/outlook/simple-types.md) del objeto `_MyEntities` (si se extrajo alguna). Cada sugerencia de tarea extraída se almacena como un objeto [TaskSuggestion](../../reference/outlook/simple-types.md) dentro de la matriz. `myGetTaskSuggestions` obtiene más información sobre cada sugerencia de tarea:


- La cadena que se identificó originalmente como una sugerencia de tarea de la propiedad [TaskSuggestion.taskString](../../reference/outlook/simple-types.md).
    
- La matriz de destinatarios de la asignación de tareas de la propiedad [TaskSuggestion.assignees](../../reference/outlook/simple-types.md). Cada destinatario de la asignación está representado en un objeto [EmailUser](../../reference/outlook/simple-types.md).
    
- Por cada persona asignada, el nombre de la propiedad [EmailUser.displayName](../../reference/outlook/simple-types.md).
    
- Por cada persona asignada, la dirección SMTP de la propiedad [EmailUser.emailAddress](../../reference/outlook/simple-types.md).
    
 `myGetTaskSuggestions` forma una cadena HTML local en `htmlText` para mostrar los datos de cada sugerencia de tarea. A continuación, se muestra el código JavaScript relacionado.




```js
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
               "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="extracting-urls"></a>Extracción de direcciones URL


Cuando el usuario hace clic en el botón  **Obtener direcciones URL**, el controlador de eventos  `myGetUrls` obtiene una matriz de direcciones URL de la propiedad [urls](../../reference/outlook/simple-types.md) del objeto `_MyEntities` (si se extrajo alguna). Cada dirección URL extraída se almacena como una cadena dentro de la matriz. `myGetUrls` forma una cadena HTML local en `htmlText` para mostrar la lista de direcciones URL extraídas.


```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="clearing-displayed-entity-strings"></a>Eliminación de las cadenas de entidades mostradas


Por último, el complemento de entidades especifica un controlador de eventos  `myClearEntitiesBox` que borra las cadenas mostradas. A continuación, se muestra el código relacionado.


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## <a name="javascript-listing"></a>Lista de JavaScript


A continuación, se muestra la lista completa de la implementación de JavaScript.


```js
// Global variables
var _Item;
var _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}


// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}

// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++)
    {
        htmlText += "E-mail Address : <span>" + 
            emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions()
{
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++)
        {
            htmlText += "Attendee : ( ";
            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + 
            meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + 
            meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + 
           meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + 
            meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
            "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
           phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="additional-resources"></a>Recursos adicionales



- [Crear complementos de Outlook para formularios de lectura](../outlook/read-scenario.md)
    
- [Coincidencia de cadenas en un elemento de Outlook como entidades conocidas](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [Método item.getEntities](../../reference/outlook/Office.context.mailbox.item.md)
    
