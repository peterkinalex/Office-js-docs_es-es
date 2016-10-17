
# <a name="validate-an-exchange-identity-token"></a>Validar un token de identidad de Exchange

Su complemento de Outlook le puede enviar un token de identidad, pero antes de que confíe en la solicitud será necesario que valide el token para asegurarse de que procede del servidor Exchange esperado. Con el ejemplo de este artículo se le mostrará cómo validar el token de identidad de Exchange con un objeto de validación escrito en C#. No obstante, para hacer la validación se puede usar cualquier lenguaje de programación. Los pasos que hay que seguir para validar el token se describen en el [borrador de Internet de token JWT (token web de JSON)](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl). 

Le recomendamos usar este procedimiento de cuatro pasos para validar el token de identidad y obtener el identificador único de usuario. En primer lugar, extraiga el token web JSON (JWT) de una cadena con codificación URL Base64. En segundo lugar, asegúrese de que el token tenga el formato correcto, que es para su complemento de Outlook, que no ha expirado y que se puede extraer una dirección URL válida para el documento de metadatos de autenticación. Después, recupere el documento de metadatos de autenticación del servidor Exchange y valide la firma adjunta al token de identidad. Por último, calcule un identificador único para el usuario. Para hacerlo, calcule el hash del id. de Exchange con la URL del documento de metadatos de autenticación. Aunque el proceso pueda parecer complejo en conjunto, cada paso individual es bastante sencillo. Puede descargar de Internet la solución que contiene estos ejemplos en [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken).
 




## <a name="set-up-to-validate-your-identity-token"></a>Configuración para validar el token de identidad


Los ejemplos de código de este artículo dependen de Windows Identity Foundation (WIF), así como un archivo DLL que extiende WIF con controladores para tokens JSON. Los ensamblados necesarios se pueden descargar desde las siguientes ubicaciones:


- [Windows Identity Foundation](http://msdn.microsoft.com/en-us/security/aa570351)
    
- [Windows.IdentityModel.Extensions.dll para aplicaciones de 32 bits](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-32.msi)
    
- [Windows.IdentityModel.Extensions.dll para aplicaciones de 64 bits](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-64.msi)
    

## <a name="extract-the-json-web-token"></a>Extracción del token web JSON


El patrón de diseño Factory Method  **Decode** divide el JWT del servidor Exchange en tres cadenas que constituyen el token, y después usa el método **Base64Decode** (mostrado en el segundo ejemplo) para descodificar el encabezado y la carga del JWT en cadenas JSON. Las cadenas se pasan al constructor **JsonToken**, donde los contenidos del JWT se validan y se devuelve una nueva instancia de objeto **JsonToken**.


```C#
    public static JsonToken Decode(string rawToken)
    {
      string[] tokenParts = rawToken.Split('.');

      if (tokenParts.Length != 3)
      {
        throw new ApplicationException("Token must have three parts separated by '.' characters.");
      }

      string encodedHeader = tokenParts[0];
      string encodedPayload = tokenParts[1];
      string signature = tokenParts[2];

      string decodedHeader = Base64Decode(encodedHeader);
      string decodedPayload = Base64Decode(encodedPayload);

      JavaScriptSerializer serializer = new JavaScriptSerializer();

      Dictionary<string, string> header = serializer.Deserialize<Dictionary<string, string>>(decodedHeader);
      Dictionary<string, string> payload = serializer.Deserialize<Dictionary<string, string>>(decodedPayload);

      return new JsonToken(header, payload, signature);
    }
```

El método **Base64Decode** implementa la lógica de descodificación que se describe en el apéndice "Notas sobre la implementación de la codificación base64url sin relleno" del[Borrador de Internet de token web JSON (JWT)](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl).




```C#
    public static Encoding TextEncoding = Encoding.UTF8;

    private static char Base64PadCharacter = '=';
    private static char Base64Character62 = '+';
    private static char Base64Character63 = '/';
    private static char Base64UrlCharacter62 = '-';
    private static char Base64UrlCharacter63 = '_';

    private static byte[] DecodeBytes(string arg)
    {
      if (String.IsNullOrEmpty(arg))
      {
        throw new ApplicationException("String to decode cannot be null or empty.");
      }

      StringBuilder s = new StringBuilder(arg);
      s.Replace(Base64UrlCharacter62, Base64Character62);
      s.Replace(Base64UrlCharacter63, Base64Character63);

      int pad = s.Length % 4;
      s.Append(Base64PadCharacter, (pad == 0) ? 0 : 4 - pad);

      return Convert.FromBase64String(s.ToString());
    }

    private static string Base64Decode(string arg)
    {
      return TextEncoding.GetString(DecodeBytes(arg));
    }
```


## <a name="parse-the-jwt"></a>Análisis del JWT


El constructor del objeto  **JsonToken** comprueba la estructura y el contenido del JWT para determinar si es válido. Es mejor hacerlo antes de solicitar el documento de metadatos de autenticación. Si el JWT no contiene las notificaciones correctas o si está fuera de su ciclo de vida, se puede evitar una llamada al servidor Exchange y el retraso asociado.

El constructor llama a métodos de utilidad para determinar si las diferentes notificaciones están presentes y son aplicables. Si se produce algún problema, el método de utilidad generará una excepción de aplicación. Si no se produce ninguna excepción, la propiedad  **IsValid** se define como **true** y el token está preparado para la validación de firma.

Cada uno de los métodos de utilidad se describe más adelante en este artículo.




```C#
    public JsonToken(Dictionary<string, string> header, Dictionary<string, string> payload, string signature)
    {

      // Assume that the token is invalid to start out.
      this.IsValid = false;

      // Set the private dictionaries that contain the claims.
      this.headerClaims = header;
      this.payloadClaims = payload;
      this.signature = signature;

      // If there is no "appctx" claim in the token, throw an ApplicationException.
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.AppContext))
      {
        throw new ApplicationException(String.Format("The {0} claim is not present.", AuthClaimTypes.AppContext));
      }

      appContext = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(payload[AuthClaimTypes.AppContext]);


      // Validate the header fields.
      this.ValidateHeader();

      // Determine whether the token is within its valid time.
      this.ValidateLifetime();

      // Validate that the token was sent to the correct URL.
      this.ValidateAudience();

      // Validate the token version.
      this.ValidateVersion();

      // Make sure that the appctx contains an authentication
      // metadata location.
      this.ValidateMetadataLocation();

      // If the token passes all the validation checks, we
      // can assume that it is valid.
      this.IsValid = true;
    }
```


### <a name="validateheader-method"></a>Método ValidateHeader

El método  **ValidateHeader** comprueba que las notificaciones necesarias están en el encabezamiento del token y que tienen los valores correctos. Es necesario configurar el encabezamiento como se indica a continuación. De lo contrario, el método generará una excepción de aplicación y se cerrará.

```js
{ "typ" : "JWT", "alg" : "RS256", "x5t" : "<thumbprint>" }
```

```C#
    private void ValidateHeaderClaim(string key, string value)
    {
      if (!this.headerClaims.ContainsKey(key))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", key));
      }

      if (!value.Equals(this.headerClaims[key]))
      {
        throw new ApplicationException(String.Format("\"{0}\" claim must be \"{0}\".", key, value));
      }
    }

    private void ValidateHeader()
    {
      ValidateHeaderClaim(AuthClaimTypes.TokenType, Config.TokenType);
      ValidateHeaderClaim(AuthClaimTypes.Algorithm, Config.Algorithm);
    
      if (!this.headerClaims.ContainsKey(AuthClaimTypes.x509Thumprint))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", AuthClaimTypes.x509Thumprint));
      }
    }


```


### <a name="validatelifetime-method"></a>Método ValidateLifetime

En el JWT se proporcionan dos fechas: "nbf" (equivalente a "no antes de" en inglés) indica la fecha y la hora en las que el token pasa a ser válido, y "exp" indica la hora a la que el token expira. Solo los tokens presentados entre estas dos fechas se deberán considerar válidos. Para dar cabida a las mínimas diferencias posibles en la configuración del reloj entre el servidor y el cliente, este método validará los tokens hasta cinco minutos antes y cinco minutos después de las horas establecidas en el token.


```C#
    private void ValidateLifetime()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidFrom))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidFrom));
      }

      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidTo))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidTo));
      }

      DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0,DateTimeKind.Utc);

      TimeSpan padding = new TimeSpan(0, 5, 0);

      DateTime validFrom = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidFrom]));
      DateTime validTo = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidTo]));

      DateTime now = DateTime.UtcNow;

      if (now < (validFrom - padding))
      {
        throw new ApplicationException(String.Format("The token is not valid until {0}.", validFrom));
      }

      if (now > (validTo + padding))
      {
        throw new ApplicationException(String.Format("The token is not valid after {0}.", validFrom));
      }
    }
```

Las fechas  **validFrom** ("nbf") y **validTo** ("exp") se envían como el número de segundos desde la época Unix, el 1 de enero de 1970. Las fechas y las horas se calculan usando el UTC para evitar problemas con las diferencias de zona horaria entre el servidor Exchange y el servidor que ejecuta el código de validación.


### <a name="validateaudience-method"></a>Método ValidateAudience

El token de identidad solo es válido para el complemento que lo solicitó. El método  **ValidateAudience** comprueba la notificación del público del token para asegurarse de que coincide con la dirección URL esperada del complemento de Outlook.


```C#
    private void ValidateAudience()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.Audience))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the application context.", AuthClaimTypes.Audience));
      }

      string location = Config.Audience.Replace("/", "-").Replace("\\", "-");
      string audience = this.payloadClaims[AuthClaimTypes.Audience].Replace("/", "-").Replace("\\", "-");

      if (!location.Equals(audience))
      {
        throw new ApplicationException(String.Format(
          "The audience URL does not match. Expected {0}; got {1}.",
          Config.Audience, this.payloadClaims[AuthClaimTypes.Audience]));
      }
    }

```


### <a name="validateversion-method"></a>Método ValidateVersion

El método  **ValidateVersion** comprueba la versión del token de identidad y se asegura de que coincide con la versión esperada. Las diferentes versiones del token pueden llevar notificaciones diferentes. La comprobación de la versión asegura que las notificaciones esperadas están en el token de identidad.


```js
    private void ValidateVersion()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchExtensionVersion))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchExtensionVersion));
      }

      if (!Config.Version.Equals(this.appContext[AuthClaimTypes.MsExchExtensionVersion]))
      {
        throw new ApplicationException(String.Format(
          "The version does not match. Expected {0}; got {1}.",
          Config.Version, this.appContext[AuthClaimTypes.MsExchExtensionVersion]));
      }
    }

```


### <a name="validatemetadatalocation-method"></a>Método ValidateMetadataLocation

El objeto de metadatos de autenticación almacenado en el servidor Exchange contiene la información necesaria para validar la firma incluida en el token de identidad. El método  **ValidateMetadataLocation** se asegura de que hay una notificación URL de metadatos de autenticación en el token de identidad al confirmar que la firma se produce en el paso siguiente.


```C#
    private void ValidateMetadataLocation()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchAuthMetadataUrl))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchAuthMetadataUrl));
      }
    }

```


## <a name="validate-the-identity-token-signature"></a>Validar la firma del token de identidad


Una vez se le ha informado de que el JWT contiene las notificaciones necesarias para validar la firma, puede usar Windows Identity Foundation (WIF) y las extensiones de WIF para validar la firma del token. Se necesitará la información siguiente para validar la firma:


- La cadena del token de identidad con codificación URL y Base 64 enviado desde el servidor Exchange.
    
- La ubicación del documento de metadatos de autenticación del JWT.
    
- La dirección URL del público del JWT.
    
En este ejemplo, el constructor de un objeto  **IdentityToken** obtiene el documento de metadatos de autenticación del servidor Exchange y valida la firma del token de identidad. Si el token de identidad es válido, se puede usar la instancia del objeto **IdentityToken** para obtener el identificador de usuario único que se incluye en el token de identidad.




```C#
    public IdentityToken(string rawToken, string audience, string authMetadataEndpoint)
    {
      X509Certificate2 currentCertificate = null;

      currentCertificate = AuthMetadata.GetSigningCertificate(new Uri(authMetadataEndpoint));

      JsonWebSecurityTokenHandler jsonTokenHandler =
          GetSecurityTokenHandler(audience, authMetadataEndpoint, currentCertificate);

      SecurityToken jsonToken = jsonTokenHandler.ReadToken(rawToken);
      JsonWebSecurityToken webToken = (JsonWebSecurityToken)jsonToken;

      SigningCertificateThumbprint = currentCertificate.Thumbprint;
      Issuer = webToken.Issuer;
      Audience = webToken.Audience;
      ValidTo = webToken.ValidTo;
      ValidFrom = webToken.ValidFrom;
      foreach (JsonWebTokenClaim claim in webToken.Claims)
      {
        if (claim.ClaimType.Equals(AuthClaimTypes.AppContextSender))
        {
          ApplicationContextSender = claim.Value;
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.IsBrowserHostedApp))
        {
          IsBrowserHostedApp = claim.Value == "true";
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.AppContext))
        {
          string[] appContextClaims = claim.Value.Split(',');
          Dictionary<string, string> appContext =
              new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(claim.Value);
          AuthenticationMetaDataUrl = appContext[AuthClaimTypes.MsExchAuthMetadataUrl];
          ExchangeID = appContext[AuthClaimTypes.MsExchImmutableId];
          TokenVersion = appContext[AuthClaimTypes.MsExchTokenVersion];
        }
      }
    }


```

La mayor parte del código del constructor del objeto  **IdentityToken** establece las propiedades de la instancia con las notificaciones del servidor Exchange. El constructor llama al método **GetSecurityTokenHandler** para obtener un controlador de tokens que validará el token de identidad de Exchange. El método **GetSecurityTokenHandler** llama a dos métodos de utilidad, **GetMetadataDocument** y **GetSigningCertificate**, que se encargan de obtener el certificado de firma desde el servidor Exchange. Cada uno de estos métodos se describe en las secciones siguientes.


### <a name="getsecuritytokenhandler-method"></a>Método GetSecurityTokenHandler

El método  **GetSecurityTokenHandler** devuelve un controlador de tokens WIF que validará el token de identidad. La mayor parte del código del método inicializa el controlador de tokens para hacer la validación. No obstante, el método llama al método **GetSigningCertificate** para recuperar el certificado X.509 usado para firmar el token del servidor Exchange.


```C#
    private JsonWebSecurityTokenHandler GetSecurityTokenHandler(string audience,
        string authMetadataEndpoint,
        X509Certificate2 currentCertificate)
    {
      JsonWebSecurityTokenHandler jsonTokenHandler = new JsonWebSecurityTokenHandler();
      jsonTokenHandler.Configuration = new SecurityTokenHandlerConfiguration();

      jsonTokenHandler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Always);
      jsonTokenHandler.Configuration.AudienceRestriction.AllowedAudienceUris.Add(
        new Uri(audience, UriKind.RelativeOrAbsolute));

      jsonTokenHandler.Configuration.CertificateValidator = X509CertificateValidator.None;

      jsonTokenHandler.Configuration.IssuerTokenResolver =
        SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
          new ReadOnlyCollection<SecurityToken>(new List<SecurityToken>(
            new SecurityToken[]
            {
              new X509SecurityToken(currentCertificate)
            })), false);

      ConfigurationBasedIssuerNameRegistry issuerNameRegistry = new ConfigurationBasedIssuerNameRegistry();
      issuerNameRegistry.AddTrustedIssuer(currentCertificate.Thumbprint, Config.ExchangeApplicationIdentifier);
      jsonTokenHandler.Configuration.IssuerNameRegistry = issuerNameRegistry;

      return jsonTokenHandler;
    }
```


### <a name="getsigningcertificate-method"></a>Método GetSigningCertificate

El método  **GetSigningCertificate** llama al método **GetMetadataDocument** para recuperar los metadatos de autenticación del servidor Exchange y devuelve el primer certificado X.509 en el documento de metadatos de autenticación. Si el documento no existe, el método genera una excepción de aplicación.


```C#
    private X509Certificate2 GetSigningCertificate(Uri authMetadataEndpoint)
    {
      JsonAuthMetadataDocument document = GetMetadataDocument(authMetadataEndpoint);

      if (null != document.keys &amp;&amp; document.keys.Length > 0)
      {
        JsonKey signingKey = document.keys[0];

        if (null != signingKey &amp;&amp; null != signingKey.keyValue)
        {
          return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
        }
      }

      throw new ApplicationException("The metadata document does not contain a signing certificate.");
    }

```


### <a name="getmetadatadocument-method"></a>Método GetMetadataDocument

El documento de metadatos de autenticación contiene la información necesaria para validar la firma en el token de identidad de Exchange. El documento se envía como una cadena JSON. El método  **GetMetatDataDocument** solicita el documento de la ubicación especificada en el token de identidad de Exchange y devuelve un objeto que encapsula la cadena JSON como un objeto. Si la dirección URL no contiene un documento de metadatos de autenticación, el método genera una excepción de aplicación.


```C#
    private JsonAuthMetadataDocument GetMetadataDocument(Uri authMetadataEndpoint)
    {
      // Uncomment the next line if your Exchange server uses the default
      // self-signed certificate.
      // ServicePointManager.ServerCertificateValidationCallback = Config.CertificateValidationCallback;

      byte[] acsMetadata;
      using (WebClient webClient = new WebClient())
      {
        acsMetadata = webClient.DownloadData(authMetadataEndpoint);
      }
      string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

      JsonAuthMetadataDocument document = new JavaScriptSerializer().Deserialize<JsonAuthMetadataDocument>(jsonResponseString);

      if (null == document)
      {
        throw new ApplicationException(String.Format("No authentication metadata document found at {0}.", authMetadataEndpoint));
      }

      return document;
    }
```

El servidor Exchange siempre usa un certificado X.509 autofirmado para autenticar las solicitudes del documento de metadatos de autenticación. Salvo que se instale un certificado que provenga de un servidor raíz, es necesario crear un método de devolución de llamada de validación del certificado. De lo contrario, la solicitud del documento de metadatos de autenticación fallará. 

La clase  **ServicePointManager** del espacio de nombres de System.Net de Framework .NET le permite vincular un método de devolución de llamada de validación configurando la propiedad **ServerCertificateValidationCallback**. Puede ver un ejemplo de un método de devolución de llamada de validación de un certificado adecuado para desarrollo y pruebas en el artículo sobre [validación de certificados X509](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx).


 **Nota de seguridad**  Si usa un método de devolución de llamada de validación de certificados, asegúrese de que cumple con los requisitos de seguridad de su organización.


## <a name="compute-the-unique-id-for-an-exchange-account"></a>Calcular el identificador exclusivo para una cuenta de Exchange


Puede crear un identificador único para una cuenta de Exchange mediante el hash de la dirección URL del documento de metadatos de autenticación con el identificador de Exchange para la cuenta. Cuando tenga este identificador único, puede usarlo para crear un sistema único de inicio de sesión (SSO) para el servicio web de su complemento de Outlook. Para más detalles sobre cómo usar el identificador único para SSO, vea [Autenticar un usuario con un token de identidad para Exchange](../outlook/authenticate-a-user-with-an-identity-token.md)

La propiedad  **UniqueUserIdentification** crea un hash de bytes aleatorios SHA256 del identificador de Exchange y de la dirección URL de los metadatos de autenticación con el proveedor SHA256 estándar del espacio de nombres **System.Security.Cryptography**.


 **Nota de seguridad**  Debe crear un hash del documento de metadatos de autenticación con el identificador de Exchange para crear un único identificador para la cuenta. Si usa solo el identificador de Exchange, puede exponer el servicio a usuarios no autorizados. Y, como siempre cuando se trata de autenticación y seguridad, debe asegurarse de que al usar el identificador único creado con este método cumple con los requisitos de seguridad de la aplicación.




```C#
    // Salt to apply when creating unique ID.
    private byte[] Salt = new byte[] {<Provide random salt bytes here };

    private string ComputeUniqueIdentification()
    {
      byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(ExchangeID, AuthenticationMetaDataUrl));

      // Combine input bytes and salt.
      byte[] saltedInput = new byte[Salt.Length + inputBytes.Length];
      Salt.CopyTo(saltedInput, 0);
      inputBytes.CopyTo(saltedInput, Salt.Length);

      // Compute the unique key.
      byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

      // Convert the hashed value to a string and return.
      return BitConverter.ToString(hashedBytes);
    }

    public string UniqueUserIdentification
    {
      get { return ComputeUniqueIdentification(); }
    }


```


## <a name="utility-objects"></a>Objetos de utilidad


Los ejemplos de código de este artículo dependen de algunos objetos de utilidad que proporcionan nombres descriptivos a las constantes que se utilizan. En la tabla siguiente se enumeran los objetos de utilidad.


**Tabla 1: Objetos de utilidad**


|**Objeto**|**Descripción**|
|:-----|:-----|
|**AuthClaimsType**|Recopila en un solo lugar los identificadores de notificación que usa el código de validación de tokens.|
|**Config**|Proporciona las constantes para validar el token de identidad. |
|**JsonAuthMetadataDocument**|Encapsula el documento de metadatos de autenticación JSON enviado desde el servidor Exchange.|

### <a name="authclaimtypes-object"></a>Objeto AuthClaimTypes

El objeto  **AuthClaimTypes** recopila en un solo lugar los identificadores de notificación que usa el código de validación de tokens. Incluye tanto las notificaciones JWT estándar como las notificaciones específicas del token de identidad de Exchange.


```C#
  public class AuthClaimTypes
  {
    public const string NameIdentifier =
        JsonWebTokenConstants.ReservedClaims.NameIdentifier;
    public const string MsExchImmutableId = "msexchuid";
    public const string MsExchTokenVersion = "version";
    public const string MsExchAuthMetadataUrl = "amurl";

    public const string AppContext =
        JsonWebTokenConstants.ReservedClaims.AppContext;
    public const string Audience =
        JsonWebTokenConstants.ReservedClaims.Audience;
    public const string Issuer =
        JsonWebTokenConstants.ReservedClaims.Issuer;
    public const string ValidFrom =
        JsonWebTokenConstants.ReservedClaims.NotBefore;
    public const string ValidTo =
        JsonWebTokenConstants.ReservedClaims.ExpiresOn;

    public const string AppContextSender = "appctxsender";
    public const string IsBrowserHostedApp = "isbrowserhostedapp";

    public const string TokenType = "typ";
    public const string Algorithm = "alg";
    public const string x509Thumbprint = "x5t";      
  }
```


### <a name="config-object"></a>Objeto Config

El objeto  **Config** contiene las constantes usadas para validar el token de identidad, además de un método de devolución de llamada de validación de certificados que puede usar si su servidor no dispone de un certificado X509 que provenga de un certificado raíz.


 
  **Nota de seguridad**  El método de devolución de llamada de certificados de seguridad solo es necesario si su servidor usa el certificado autofirmado predeterminado. El método de devolución de llamada de este ejemplo devuelve  **false** cuando el certificado es autofirmado, de manera que será necesario cambiarlo por un método de devolución de llamada que cumpla con los requisitos de seguridad de su organización. Para ver un ejemplo de un método de devolución de llamada de validación de certificados que sea adecuado para desarrollo y pruebas, visite [Validación de certificados X509](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx).


```C#
  public static class Config
  {
    public static string Algorithm = "RS256";
    public static string Audience = @"https:\\localhost:44300\Pages\IdentityTest.html";
    public static string TokenType = "JWT";
    public static string Version = "ExIdTok.V1";

    public static string ExchangeApplicationIdentifier = "Exchange";

    internal static bool CertificateValidationCallback(
    object sender,
    System.Security.Cryptography.X509Certificates.X509Certificate certificate,
    System.Security.Cryptography.X509Certificates.X509Chain chain,
    System.Net.Security.SslPolicyErrors sslPolicyErrors)
    {
      // If the certificate is a valid, signed certificate, return true.
      if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
      {
        return true;
      }

      // If there are errors in the certificate chain, look at each error to determine the cause.
      else
      {
        return false;
      }
    }
  }
```


### <a name="jsonauthmetadatadocument-object"></a>Objeto JsonAuthMetadataDocument

El objeto  **JsonAuthMetadataDocument** expone los contenidos del documento de metadatos de autenticación con las propiedades.


```C#
using System;

namespace IdentityTest
{
  public class JsonAuthMetadataDocument
  {
    public string id { get; set; }
    public string version { get; set; }
    public string name { get; set; }
    public string realm { get; set; }
    public string serviceName { get; set; }
    public string issuer { get; set; }
    public string [] allowedAudiences { get; set; }
    public JsonKey[] keys;
    public JsonEndpoint[] endpoints;
  }

  public class JsonEndpoint
  {
    public string location { get; set; }
    public string protocol { get; set; }
    public string usage { get; set; }
  }

  public class JsonKey
  {
    public string usage { get; set; }
    public JsonKeyValue keyValue { get; set; }
  }

  public class JsonKeyValue
  {
    public string type { get; set; }
    public string value { get; set; }
  }
}

```


## <a name="additional-resources"></a>Recursos adicionales



- [Autenticar un complemento de Outlook con los tokens de identidad de Exchange](../outlook/authentication.md)
    
- [Contenido del token de identidad de Exchange](../outlook/inside-the-identity-token.md)
    
