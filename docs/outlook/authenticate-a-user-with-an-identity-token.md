
# Autenticar un usuario con un token de identidad para Exchange

Puede implementar un esquema de autenticación de inicio de sesión único (SSO) para un servicio de información que permite que los clientes que usan complementos de Outlook se conecten con su servicio a través de las credenciales del servidor Exchange. En este artículo se explica cómo hacer coincidir las credenciales con un almacén de datos de usuario simple basado en objetos  **Dictionary**.

 >**Nota**  Este es solo un ejemplo de SSO y no se debe usar en el código de producción. Como siempre, cuando esté trabajando con identidades y autenticaciones, deberá asegurarse de que el código cumpla los requisitos de seguridad de su organización.


## Requisitos previos para usar la autenticación SSO


Para usar un token de identidad para SSO, la aplicación de servicio debe tener un token de identidad válido. Puede obtener información sobre los tokens de identidad y cómo solicitar y validar un token de identidad en los siguientes artículos:


- [Contenido del token de identidad de Exchange](../outlook/inside-the-identity-token.md)
    
- [Llamar a un servicio de un complemento de Outlook con un token de identidad en Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Usar la biblioteca de validación de tokens de Exchange](../outlook/use-the-token-validation-library.md) si usa código administrado, o [Validar un token de identidad de Exchange](../outlook/validate-an-identity-token.md) si escribe su propio método de validación de tokens.
    

## Autenticación de usuarios


El siguiente ejemplo de código muestra un objeto de autenticación simple que empareja la identidad única representada por un token de identidad con un conjunto de credenciales para un servicio. La clase  **TokenAuthentication** proporciona un método, **GetResponseFromService**, que devolverá una respuesta para los tokens previamente autenticados, o bien solicitará al usuario que proporcione credenciales que puedan ser autenticadas y asociadas con el token de identidad. El código no está completo, se supone que se proporcionarán los siguientes objetos y métodos.



|**Objeto/método**|**Descripción**|
|:-----|:-----|
|Objeto **LocalCredentials**|Representa las credenciales del usuario para el servicio. La estructura del objeto depende de los requisitos del servicio.|
|Objeto **IdentityToken**|Contiene un token de identidad de usuario enviado al servicio por un complemento de Outlook. El objeto debe contener al menos el identificador único de Exchange del usuario y la dirección URL de metadatos de autenticación para el servidor que emitió el token. En este ejemplo, se usa el objeto de token de identidad que se define en el artículo [Validar un token de identidad de Exchange](../outlook/validate-an-identity-token.md).|
|Objeto **JsonResponse**|Representa la respuesta del servicio. El objeto se puede serializar a un objeto JSON.|
|Método **CallService**|Llama al servicio con un objeto  **LocalCredentials** que contiene las credenciales de usuario para el servicio y un objeto que contiene datos para la solicitud del servicio. Si las credenciales son válidas, este método devuelve un objeto **JsonReponse** que contiene los resultados de la solicitud. Si las credenciales no son válidas, este método devuelve **null**.|
|Método **GetCredentialsResponse**|Devuelve un objeto  **JsonReponse** que el complemento de correo de Office reconocerá como una solicitud de credenciales para el servicio.|
|Método **LocalCredentialsAreValid**|Devuelve  **true** si las credenciales provistas al servicio son válidas; de lo contrario, devuelve **false**.|

 >**Nota**  Esto es una sugerencia sobre cómo usar el token de identidad. Como siempre, cuando esté trabajando con identidades y autenticaciones, deberá asegurarse de que el código cumpla los requisitos de seguridad de su organización.


```C#
    public class TokenAuthentication
    {
        // This example uses a Dictionary object to store local credentials. Your application should use
        // a data store that is appropriate to the security requirements of your organization.
        private Dictionary<string, LocalCredentials> AuthenticationCache = new Dictionary<string, LocalCredentials>();

        // Salt to apply when creating unique ID.
        private byte[] Salt = new byte[] {25, 139, 201, 13};

        private JsonResponse CallService(LocalCredentials credentials, object data)
        {
            // Calls the local service to get the response for the user.
            return null;
        }

        private JsonResponse GetCredentialsResponse()
        {
            // Creates a response that tells the Outlook add-in to
            // request the user's credentials for the service.
            return null;
        }

        private bool LocalCredentialsAreValid(LocalCredentials credentials)
        {
            // Returns true if the service recognizes the credentials provided.
            return false;
        }

        private string ComputeSHA256Hash(string uniqueId, string authenticationMetadataUrl, byte[] salt)
        {
            byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(uniqueId, authenticationMetadataUrl));

            // Combine input bytes and salt.
            byte[] saltedInput = new byte[salt.Length + inputBytes.Length];
            salt.CopyTo(saltedInput, 0);
            inputBytes.CopyTo(saltedInput, salt.Length);

            // Compute the unique key.
            byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

            // Convert the hashed value to a string and return.
            return BitConverter.ToString(hashedBytes);
        }

        public JsonResponse GetResponseFromService(IdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // The user's credentials are in the cache; make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials.
                    string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
    }}
```


## Autenticación de un usuario con la biblioteca de validación administrada


Si usa la biblioteca administrada para validar tokens de identidad, no es necesario que aplique una clave única. La propiedad  **UniqueUserIdentification** en la clase **AppIdentityToken** se puede usar directamente como clave única para el usuario. En el siguiente ejemplo de código se muestran las modificaciones al método **GetResponseFromService** en el ejemplo anterior que debe hacer para usar la clase **AppIdentityToken**.


```js
        public JsonResponse GetResponseFromService(AppIdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = token.UniqueUserIdentitification;
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // User's credentials are in the cache. Make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials. 
                    string uniqueKey = token.UniqueUserIdentitification;
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
```


## Recursos adicionales



- [Autenticar un complemento de Outlook con los tokens de identidad de Exchange](../outlook/authentication.md)
    
- [Llamar a un servicio de un complemento de Outlook con un token de identidad en Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Usar la biblioteca de validación de tokens de Exchange](../outlook/use-the-token-validation-library.md)
    
- [Validar un token de identidad de Exchange](../outlook/validate-an-identity-token.md)
    
