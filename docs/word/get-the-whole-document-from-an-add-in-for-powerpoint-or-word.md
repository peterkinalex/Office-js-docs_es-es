
# <a name="get-the-whole-document-from-an-add-in-for-powerpoint-or-word"></a>Procedimiento para obtener el documento completo de un complemento para PowerPoint o Word

Puede crear un Complemento de Office para permitir enviar o publicar con un clic un documento de Word 2013 o PowerPoint 2013 en una ubicación remota. En este artículo se muestra cómo comopilar un complemento de panel de tareas sencillo para PowerPoint 2013 que obtenga toda la presentación como un objeto de datos y envíe esos datos a un servidor web mediante una solicitud HTTP.

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a>Requisitos previos para crear un complemento para PowerPoint o Word


En este artículo se supone que se usa un editor de texto para crear el complemento de panel de tareas para PowerPoint o Word. Para crear el complemento de panel de tareas, debe crear los siguientes archivos:


- En una carpeta de red compartida o en un servidor web, necesita los siguientes archivos:
    
      - Un archivo HTML (GetDoc_App.html) que contenga la interfaz de usuario junto con vínculos a los archivos JavaScript (incluidos office.js y host-specific .js y los archivos de hoja de estilos en cascada (CSS).
    
  - Un archivo JavaScript (GetDoc_App.js) que contenga la lógica de programación del complemento.
    
  - Un archivo CSS (Program.css) que contenga los estilos y los formatos del complemento.
    
- Un archivo de manifiesto XML (GetDoc_App.xml) para el complemento, disponible en un catálogo de complementos o una carpeta de red compartida. El archivo de manifiesto debe apuntar a la ubicación del archivo HTML mencionado anteriormente.
    
También puede crear un complemento para PowerPoint o Word con Visual Studio 2015. 


### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a>Conceptos básicos que deben considerarse al crear un complemento de panel de tareas

Antes de empezar a crear este complemento para PowerPoint o Word, debe estar familiarizado con la compilación de Complementos de Office y estar acostumbrado a trabajar con solicitudes HTTP. En este artículo no se explica cómo descodificar texto con codificación Base64 de una solicitud HTTP en un servidor web. 


## <a name="create-the-manifest-for-the-add-in"></a>Crear el manifiesto del complemento


El archivo de manifiesto XML del complemento para PowerPoint ofrece información importante sobre el complemento: qué aplicaciones pueden hospedarlo, la ubicación del archivo HTML, el título y la descripción del complemento, y muchas otras características.


- En un editor de texto, agregue el siguiente código al archivo de manifiesto.
    
```XML
  
<?xml version="1.0" encoding="utf-8" ?> 
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="TaskPaneApp">
    <Id>[Replace_With_Your_GUID]</Id> 
    <Version>1.0</Version> 
    <ProviderName>[Provider Name]</ProviderName> 
    <DefaultLocale>EN-US</DefaultLocale> 
    <DisplayName DefaultValue="Get Doc add-in" /> 
    <Description DefaultValue="My get PowerPoint or Word document add-in." /> 
    <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" /> 
    <Hosts>
      <Host Name="Document" /> 
      <Host Name="Presentation" /> 
    </Hosts>
    <DefaultSettings>
      <SourceLocation DefaultValue="[Network location of app]/GetDoc_App.html" /> 
    </DefaultSettings>
    <Permissions>ReadWriteDocument</Permissions> 
</OfficeApp>
```

- Guarde el archivo como GetDoc_App.xml con codificación UTF-8 en una ubicación de red o en un catálogo de complementos.
    

## <a name="create-the-user-interface-for-the-add-in"></a>Crear la interfaz de usuario del complemento


Para la interfaz de usuario del complemento, puede usar HTML, escrito directamente en el archivo GetDoc_App.html. La lógica de programación y la funcionalidad del complemento deben estar en un archivo JavaScript (por ejemplo, GetDoc_App.js).

Use el siguiente procedimiento para crear una interfaz de usuario simple para el complemento, que incluya un título y un solo botón.


- En el editor de texto, en un archivo nuevo, agregue el siguiente código HTML.
    
```html  
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
        <title>Publish presentation</title>
        <link rel="stylesheet" type="text/css" href="Program.css" />
        <script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
        <script src="GetDoc_App.js"></script>
    </head>
    <body>
      <form>
        <h1>Publish presentation</h1>
        <br />
        <div><input id='submit' type="button" value="Submit" /></div>
        <br />
        <div><h2>Status</h2> 
            <div id="status"></div>
        </div>
      </form>
    </body>
</html>
```

- Guarde el archivo como GetDoc_App.html con codificación UTF-8 en una ubicación de red o en un servidor web.
    

 >**Nota** Asegúrese de que las etiquetas **head** del complemento contengan una etiqueta de **script** con un vínculo válido al archivo office.js. 

Usaremos un poco de CSS para que el complemento tenga una apariencia sencilla y, al mismo tiempo, moderna y profesional. Use el siguiente código CSS para definir el estilo del complemento.


- En el editor de texto, en un archivo nuevo, agregue el siguiente código CSS.
    
```css 
body
{
    font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
}
h1,h2
{
    text-decoration-color:#4ec724;
}
input [type="submit"], input[type="button"] 
{ 
    height:24px; 
    padding-left:1em; 
    padding-right:1em; 
    background-color:white; 
    border:1px solid grey; 
    border-color: #dedfe0 #b9b9b9 #b9b9b9 #dedfe0; 
    cursor:pointer; 
}
```

- Guarde el archivo como Program.css con codificación UTF-8 en la ubicación de red o en el servidor web donde se encuentra el archivo GetDoc_App.html.
    

## <a name="add-the-javascript-to-get-the-document"></a>Agregar el código JavaScript para obtener el documento


En el código del complemento, un controlador del evento [Office.initialize](../../reference/shared/office.initialize.md) agrega un controlador al evento de clic del botón **Enviar** en el formulario e informa al usuario de que el complemento está listo.

En el siguiente ejemplo de código se muestra el controlador de eventos del evento  **Office.initialize** junto con una función auxiliar, `updateStatus`, para escribir en el div de estado.




```js
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {

      // After the DOM is loaded, add-in-specific code can run.
      document.getElementById('submit').addEventListener("click",
          function () {
              sendFile();
          });}
      updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div. 
function updateStatus(message) {
    var statusInfo = document.getElementById("status");
    statusInfo.innerHTML += message + "<br/>";
}
```



Al elegir el botón **Enviar** en la interfaz de usuario, el complemento llama a la función `sendFile`, que contiene una llamada al método [Document.getFileAsync](../../reference/shared/document.getfileasync.md). El método **getFileAsync** usa un modelo asincrónico, igual que otros métodos de la API de JavaScript para Office. Tiene un parámetro obligatorio, _fileType_, y dos parámetros opcionales, _options_ y _callback_. 

El parámetro  _fileType_ espera una de las tres constantes de la enumeración [FileType](../../reference/shared/filetype-enumeration.md):  **Office.FileType.Compressed** ("comprimido"), **Office.FileType.PDF** ("pdf") o **Office.FileType.Text** ("texto"). PowerPoint solo es compatible con **Compressed** como argumento; Word es compatible con las tres. Al pasar **Compressed** para el parámetro _fileType_, el método  **getFileAsync** devuelve el documento como un archivo de presentación de PowerPoint 2013 (*.pptx) o archivo de documento de Word 2013 (*.docx) al crear una copia temporal del archivo en el equipo local.

El método **getFileAsync** devuelve una referencia al archivo como un objeto [File](../../reference/shared/file.md). El objeto **File** expone cuatro miembros: la propiedad [size](../../reference/shared/file.size.md), la propiedad [sliceCount](../../reference/shared/file.slicecount.md), el método [getSliceAsync](../../reference/shared/file.getsliceasync.md) y el método [closeAsync](../../reference/shared/file.closeasync.md). La propiedad **size** devuelve el número de bytes que tiene el archivo. La propiedad **sliceCount** devuelve el número de objetos [Slice](../../reference/shared/document.md) (se explica posteriormente en este artículo) del archivo.

El siguiente código recupera un documento de Word o de PowerPoint como un objeto **File** mediante el método **document.getFileAsync()**. Después, empaqueta el objeto **File** resultante, un contador a cero y [sliceCount](../../reference/shared/file.slicecount.md) en un objeto anónimo. Este objeto se pasa posteriormente a una función `getSlice` definida localmente. 

```js
// Get all the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {

    Office.context.document.getFileAsync("compressed",
        { sliceSize: 100000 },
        function (result) {

            if (result.status == Office.AsyncResultStatus.Succeeded) {

                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size +
                    " bytes");

                getSlice(state);
            }
            else {
                updateStatus(result.status);
            }
    });
}
```

La función local  `getSlice` realiza una llamada al método **File.getSliceAsync** para recuperar un segmento del objeto **File**. El método  **getSliceAsync** devuelve un objeto **Slice** de la colección de segmentos. Tiene dos parámetros obligatorios: _sliceIndex_ y _callback_. El parámetro  _sliceIndex_ toma un entero como indizador en la colección de segmentos. Al igual que otras funciones de la API de JavaScript para Office, el método **getSliceAsync** también toma una función de devolución de llamada como parámetro para controlar los resultados de la llamada al método.

El objeto **Slice** le permite obtener acceso a los datos que contiene el archivo. Excepto si se especifica lo contrario en el parámetro _options_ del método **getFileAsync**, el objeto **Slice** tiene un tamaño de 4 MB. El objeto **Slice** expone tres propiedades: [size](../../reference/shared/slice.size.md), [data](../../reference/shared/slice.data.md) e [index](../../reference/shared/slice.index.md). La propiedad **size** obtiene el tamaño, en bytes, del segmento. La propiedad **index** obtiene un entero que representa la posición del segmento dentro de la colección de segmentos.




```js

// Get a slice from the file and then call sendSlice.
function getSlice(state) {

    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {

            updateStatus("Sending piece " + (state.counter + 1) +
                " of " + state.sliceCount);

            sendSlice(result.value, state);
        }
        else {
            updateStatus(result.status);
        }
    });
}
```

La propiedad  **Slice.data** devuelve los datos sin procesar del archivo como una matriz de bytes. Si los datos están en formato de texto (es decir, XML o texto sin formato), el segmento contiene el texto sin formato. Si pasa **Office.FileType.Compressed** para el parámetro _fileType_ de **Document.getFileAsync**, el segmento contiene los datos binarios del archivo como una matriz de bytes. En el caso de los archivos de PowerPoint o Word, los segmentos contienen matrices de bytes.

Debe implementar su propia función (o usar una biblioteca disponible) para convertir los datos de la matriz de bytes en una cadena codificada en Base64. Para obtener más información sobre la codificación en Base64 con JavaScript, consulte [Codificación y descodificación en Base64](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).

Después de convertir los datos a Base64, puede transmitirlos a un servidor web de varias maneras (por ejemplo, como el cuerpo de una solicitud HTTP POST).

Agregue el siguiente código para enviar un segmento a un servicio web.


 >**Nota**  Este código envía un archivo de PowerPoint o Wordal servidor web en varios segmentos. El servidor web o el servicio debe compilar todos los segmentos en un solo archivo .pptx para que pueda manipularlo.




```js

function sendSlice(slice, state) {
    var data = slice.data;

    // If the slice contains data, create an HTTP request.
    if (data) {

        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't 
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/en-US/docs/Web/JavaScript/Base64_encoding_and_decoding.
        var fileData = myEncodeBase64(data);

        // Create a new HTTP request. You need to send the request 
        // to a webpage that can receive a post.
        var request = new XMLHttpRequest();

        // Create a handler function to update the status 
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                }
                else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "[Your receiving page or service]");
        request.setRequestHeader("Slice-Number", slice.index);

        // Send the file as the body of an HTTP POST 
        // request to the web server.
        request.send(fileData);
    }
}
```



Como su nombre indica, el método  **File.closeAsync** cierra la conexión al documento y libera recursos. Aunque el recolector de elementos no utilizados del espacio aislado de los Complementos de Office recopila las referencias a archivos que se encuentren fuera del ámbito, sigue siendo recomendable cerrar expresamente los archivos cuando el código haya terminado de trabajar con ellos. El método **closeAsync** tiene un solo parámetro, _callback_, que especifica la función de llamada al finalizar la llamada.




```js

function closeFile(state) {

    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {

        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status == "succeeded") {
            updateStatus("File closed.");
        }
        else {
            updateStatus("File couldn't be closed.");
        }
    });
}
```
