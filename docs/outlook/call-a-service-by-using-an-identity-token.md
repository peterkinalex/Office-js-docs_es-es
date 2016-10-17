
# <a name="call-a-service-from-an-outlook-add-in-by-using-an-identity-token-in-exchange"></a>Llamar a un servicio de un complemento de Outlook con un token de identidad en Exchange

Un token de identidad proporciona un identificador único para cada uno de sus clientes que puede usar para personalizar el servicio que ofrece. El código puede pedir un token de identidad al servidor Exchange mediante una llamada de método asincrónico que devuelve una cadena al complemento de Outlook. Esa cadena contiene un token de identidad JWT (por token web de JSON). El complemento no necesita desempaquetar el token, lo remite a su servicio web para que este pueda autenticar la solicitud del complemento.

El servicio web compatible con su complemento debe ejecutarse en el mismo servidor que hospeda los archivos de origen HTML y JavaScript del complemento. De este modo, se evitan errores de scripting entre los sitios. Su servidor puede remitir la solicitud a otros servicios web si la aplicación así lo pide.

Agregar un token de identidad a la solicitud de servicio que envía el complemento es fácil: solicite el token, úselo y, después, use la respuesta del servicio web. Aquí tiene un ejemplo de un sencillo documento XML enviado a su servidor con el método **XmlHttpRequest**.

## <a name="request-a-token-from-your-exchange-server"></a>Pedir un token al servidor Exchange


Este sencillo método de inicialización para un complemento usa el método  **getUserIdentityTokenAsync** para pedir un token de identidad del servidor Exchange. El parámetro _getUserIdentityToken_ es la función que se llama cuando se devuelve la solicitud asincrónica al servidor. En el siguiente paso se explica el método de devolución de llamada.


```js
var _mailbox;
var _xhr;
// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
        _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}

```


## <a name="use-the-identity-token"></a>Usar el token de identidad


La función de devolución de llamada para el método  **getUserIdentityTokenAsync** tiene un parámetro que contiene el token de identidad del usuario en su propiedad **value**.

Esta función de devolución de llamada crea un objeto  **XMLHttpRequest** para llamar al servicio web. Configure la propiedad **onreadystatechange** del objeto **XMLHttpRequest** con el nombre de la función que debe ejecutarse cuando el complemento recibe una respuesta del servicio web.




```js
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}
```


## <a name="use-the-web-service-response"></a>Usar la respuesta del servicio web


Esta es otra función simple que procesa la respuesta del servicio web. Sigue el patrón estándar para las funciones de devolución de llamada de  **XHMHttpResponse**. Espera a que llegue la respuesta entera del servicio web y entonces coloca el contenido de la respuesta en la interfaz de usuario del complemento. La respuesta que esa función analiza es la respuesta del servicio web. Para más información sobre esta respuesta, vea [Validar un token de identidad de Exchange](../outlook/validate-an-identity-token.md). 


```js
function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


## <a name="example:-calling-a-web-service-with-identity-tokens"></a>Ejemplo: llamar a un servicio web con tokens de identidad


Los tokens de identidad proporcionan información sobre la identidad del cliente que llama al servicio a un servicio web que se ejecuta en el servidor. Para usar tokens de identidad, necesitará lo siguiente:


- Un complemento de Outlook que pida un token de identidad al servidor Exchange y lo envíe al servicio web. La información contenida en este tema lo ayudará a crear dicho complemento.
    
- Un servicio web que se ejecute en el servidor y que proporcione la interfaz de usuario del complemento que valida el token de identidad. Encontrará la información que necesita para crear el servicio web en los temas siguientes:
    
      - [Usar la biblioteca de validación de tokens de Exchange:](../outlook/use-the-token-validation-library.md) si usa la biblioteca de validación que proporcionamos.
    
  - [Validar un token de identidad de Exchange:](../outlook/validate-an-identity-token.md) si escribe su propio código de validación.
    

### <a name="code-for-the-sample-add-in"></a>Código para el complemento de ejemplo


Los siguientes archivos son necesarios para el complemento que se describe en este artículo:


- IdentityTest.js: los archivos de JavaScript que proporcionan la lógica de negocios para el complemento.
    
- IdentityTest.html: el archivo HTML que proporciona la interfaz de usuario para el complemento.
    
También necesitará el servicio web de prueba de identidad. Para más información sobre este servicio web, vea [Validar un token de identidad de Exchange](../outlook/validate-an-identity-token.md).


#### <a name="identitytest.js"></a>IdentityTest.js

En el siguiente ejemplo se muestra el archivo IdentityTest.js.


```js
var _mailbox;
var _xhr;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}

function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


#### <a name="identitytest.html"></a>IdentityTest.html

En el siguiente ejemplo se muestra el archivo IdentityTest.html.


```HTML
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Identity Test</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <script src="../Scripts/jquery-1.6.2.js"></script>
    <script src="../Scripts/Office/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/Office.js"></script>

    <!-- Add your JavaScript to the following JavaScript file -->
    <script src="../Scripts/IdentityTest.js"></script>
</head>
<body>
    <div id="SectionContent">
        <table style="width: 80%;">
            <tr>
                <th>Claim
                </th>
                <th>Contents
                </th>
            </tr>
            <tr>
                <td style="width: 25%;">Error:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="error" value="None" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">User Exchange ID:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="msexchuid" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Authentication Metadata URL:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="amurl" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Unique identifier:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="uniqueID" />
                </td>
            </tr>
          </tr>
            <tr>
                <td style="width: 25%;">Audience:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="aud" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Issuer:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="iss" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Certificate thumbprint:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="x5t" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid from:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="nbf" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid to:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="exp" />
                </td>
            </tr>
        </table>
    </div>
</body>
</html>
```


## <a name="next-steps"></a>Pasos siguientes


Ahora que sabe cómo solicitar un token de identidad, debe usar el token en el lado del servidor de la solicitud. Los siguientes artículos lo ayudarán a comenzar:


- [Usar la biblioteca de validación de tokens de Exchange](../outlook/use-the-token-validation-library.md)
    
- [Validar un token de identidad de Exchange](../outlook/validate-an-identity-token.md)
    
- [Autenticar un usuario con un token de identidad para Exchange](../outlook/authenticate-a-user-with-an-identity-token.md)
    

## <a name="additional-resources"></a>Recursos adicionales



- [Autenticar un complemento de Outlook con los tokens de identidad de Exchange](../outlook/authentication.md)
    
- [Contenido del token de identidad de Exchange](../outlook/inside-the-identity-token.md)
    
