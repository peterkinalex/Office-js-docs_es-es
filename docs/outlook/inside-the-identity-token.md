
# <a name="inside-the-exchange-identity-token"></a>Contenido del token de identidad de Exchange
Descubra qué hay dentro de un token de identidad de Exchange 2013.



El token de identidad de autenticación que envía el servidor Exchange al complemento de Outlook es opaco para el complemento. No debe examinar el contenido del token para enviarlo al servidor, pero cuando escriba el código del servicio web que interactúa con el complemento de Outlook, necesitará saber qué hay dentro del token de identidad.

## <a name="what-is-an-identity-token"></a>¿Qué es un token de identidad?


Un token de identidad es una cadena con codificación URL base64 que está autofirmada por el servidor Exchange que la envió. El token no está cifrado, y la clave pública que usa para validar la firma se almacena en el servidor Exchange que emitió el token. El token tiene tres partes: un encabezado, una carga y una firma. En la cadena del token, las partes están separadas por el carácter "." para facilitar la tarea de división del token.

Exchange 2013 usa un token web JSON (JWT) para el token de identidad. Para obtener información sobre tokens JWT, consulte el [borrador de Internet del token web JSON (JWT)](http://self-issued.info/docs/draft-ietf-oauth-json-web-token.html).


### <a name="identity-token-header"></a>Encabezado del token de identidad

El encabezado identifica el token y permite al servicio web saber cuál es el tipo de token que se presenta. En el siguiente ejemplo se muestra el aspecto del encabezado del token.

```js
{ "typ" : "JWT", "alg" : "RS256", "x5t" : "Un6V7lYN-rMgaCoFSTO5z707X-4" }
```

En la siguiente tabla se describen las partes del encabezado del token de identidad.


**Partes del encabezado del token de identidad**


|**Notificación**|**Valor**|**Descripción**|
|:-----|:-----|:-----|
|typ|"JWT"|Identifica el token como un token web JSON. Todos los tokens de identidad proporcionados por el servidor Exchange son tokens JWT.|
|alg|"RS256"|El algoritmo hash que se usa para crear la firma. Todos los tokens proporcionados por el servidor Exchange usan el algoritmo RS-256.|
|x5t|Huella digital de certificado|La huella digital X.509 del token.|

### <a name="identity-token-payload"></a>Carga del token de identidad

La carga contiene las notificaciones de autenticación que identifican la cuenta de correo electrónico y el servidor Exchange que envió el token. En el siguiente ejemplo se muestra el aspecto de la sección de carga.
```js

{ 
   "aud" : "https://mailhost.contoso.com/IdentityTest.html", 
   "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", 
   "nbf" : "1331579055", 
   "exp" : "1331607855", 
   "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
   "isbrowserhostedapp":"true",
"appctx" : { 
     "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com" "version" : "ExIdTok.V1" "amurl" :         "https://mailhost.contoso.com:443/autodiscover/metadata/json/1" 
     } 
}
```
En la siguiente tabla se enumeran las partes de la carga del token de identidad.


**Partes de la carga del token de identidad**


|**Notificación**|**Descripción**|
|:-----|:-----|
|aud|La dirección URL del complemento que solicitó el token. Un token solo es válido si se envía desde el complemento que se ejecuta en el explorador del cliente. Si el complemento usa el esquema de manifiestos v1.1 de Complementos de Office, esta URL es la especificada en el primer elemento  **SourceLocation** en el tipo de formulario **ItemRead** o **ItemEdit**, el que aparezca primero como parte del elemento [FormSettings](http://msdn.microsoft.com/en-us/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx) en el manifiesto del complemento.|
|iss|Un identificador único del servidor Exchange que emitió el token. Todos los tokens emitidos por este servidor Exchange tendrán el mismo identificador.|
|nbf|La fecha y la hora de inicio de la validez del token. El valor es el número de segundos desde el 1 de enero de 1970. |
|exp|La fecha y la hora de finalización de la validez del token. El valor es el número de segundos desde el 1 de enero de 1970.|
|appctxsender|Identificador único para el servidor Exchange que envió el contexto de la aplicación.|
|isbrowserhostedapp|Indica si el complemento se hospeda en un explorador.|
|appctx|El contexto de aplicación del token. |
La información de la notificación appctx proporciona la dirección de la cuenta de correo electrónico y un identificador único de la cuenta. En la siguiente tabla se enumeran las partes de la notificación appctx.



|**Parte de la notificación appctx**|**Descripción**|
|:-----|:-----|
|msexchuid|Un identificador único asociado a la cuenta de correo electrónico y el servidor Exchange.|
|version|El número de versión del token. En todos los tokens proporcionados por un servidor que ejecuta Exchange 2013, el valor es "ExIdTok.V1".|
|amurl|La dirección URL del documento de metadatos de autenticación que contiene la clave pública del certificado X.509 que se usó para firmar el token. Para más información sobre cómo usar el documento de metadatos de autenticación, vea [Validar un token de identidad de Exchange](../outlook/validate-an-identity-token.md).|

### <a name="identity-token-signature"></a>Firma del token de identidad

La firma se crea al cifrar las secciones de encabezado y de carga con el algoritmo hash especificado en el encabezado y con el certificado X.509 autofirmado que se encuentra en el servidor en la ubicación especificada en la carga. El servicio web puede validar esta firma para garantizar que el token de identidad proviene del servidor que se espera que lo envíe.


## <a name="additional-resources"></a>Recursos adicionales



- [Autenticar un complemento de Outlook con los tokens de identidad de Exchange](../outlook/authentication.md)
    
- [Llamar a un servicio de un complemento de Outlook con un token de identidad en Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Usar la biblioteca de validación de tokens de Exchange](../outlook/use-the-token-validation-library.md)
    
- [Validar un token de identidad de Exchange](../outlook/validate-an-identity-token.md)
    
