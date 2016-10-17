
# <a name="use-the-exchange-web-services-managed-api-token-validation-library"></a>Usar la biblioteca de validación de tokens de API administrados de servicios Web Exchange

Puede identificar a los clientes de su complemento de Outlook con un token de identidad que su complemento solicite de un servidor que ejecute Exchange Server 2013 o Exchange Online. El token, con el formato de token web JSON, proporciona un identificador único de cuenta de correo electrónico en un servidor Exchange. La API administrada de servicios Web Exchange (EWS) proporciona clases auxiliares que simplifican el uso del token de identidad.

## <a name="prerequisites-for-using-the-validation-library"></a>Requisitos previos para el uso de bibliotecas de validación

Para validar un token de identidad de Exchange, tiene que instalar la [biblioteca de API administrada de EWS](https://www.nuget.org/packages/Microsoft.Exchange.WebServices).

## <a name="validate-the-exchange-identity-token"></a>Validación del token de identidad de Exchange

La biblioteca de validación de la API administrada con EWS proporciona la clase **AppIdentityToken** para administrar los tokens de identidad de Exchange. En el método siguiente se muestra cómo crear una instancia de **AppIdentityToken** y llamar al método **Validate** para comprobar que el token es válido. El método admite los parámetros siguientes:

- *rawToken*: La representación de cadena del token devuelto en el complemento de Outlook desde el método [**Office.context.mailbox.getUserIdentityTokenAsync**](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox).
- *hostUri*: El URI completo a la página del complemento de Outlook que se denomina **getUserIdentityTokenAsync**.

```C#
// Required to use the validation library.
using Microsoft.Exchange.WebServices.Auth.Validate;

private AppIdentityToken CreateAndValidateIdentityToken(string rawToken, string hostUri)
{
    try
    {
        AppIdentityToken token = (AppIdentityToken)AuthToken.Parse(rawToken);
        token.Validate(new Uri(hostUri));

        return token;
    }
    catch (TokenValidationException ex)
    {
        throw new ApplicationException("A client identity token validation error occurred.", ex);
    }
}
```

## <a name="additional-resources"></a>Recursos adicionales

- [Autenticar un complemento de Outlook con los tokens de identidad de Exchange](../outlook/authentication.md)  
- [Contenido del token de identidad de Exchange](../outlook/inside-the-identity-token.md)
- [Validar un token de identidad de Exchange](../outlook/validate-an-identity-token.md)
    
