# Autorizar servicios externos en el complemento de Office

Los servicios en línea más populares, como Office 365, Google, Facebook, LinkedIn, SalesForce y GitHub, permiten a los desarrolladores proporcionar a los usuarios acceso a sus cuentas en otras aplicaciones. Esto le abre la posibilidad de incluir estos servicios en su complemento de Office. 

El marco de trabajo estándar de la industria para habilitar el acceso de una aplicación web a un servicio en línea se llama OAuth 2.0. En la mayoría de los casos, no es necesario conocer los detalles de cómo funciona el marco para hacer uso de él en el complemento. Existen muchas bibliotecas que abstraen los detalles.

Una idea fundamental de OAuth es que una aplicación puede ser una entidad de seguridad ante sí misma, al igual que un usuario o un grupo, con su propia identidad y un conjunto de permisos. En los casos más típicos, cuando el usuario realiza una acción en el complemento de Office que solicita el servicio en línea, el complemento envía al servicio una solicitud para un conjunto específico de permisos en la cuenta del usuario. Después, el servicio pide al usuario que conceda esos permisos al complemento. Cuando se conceden los permisos, el servicio envía al complemento un pequeño *token de acceso* codificado. Para usar el servicio, el complemento incluye el token en todas sus solicitudes a las API del servicio. Pero el complemento solo puede actuar con los permisos que le ha concedido el usuario. Además, el token caduca después del tiempo especificado.

Hay varios modelos de OAuth, llamados *flujos* o *tipos de concesión*, diseñados para distintos escenarios. Los siguientes son los dos más importantes:

- **Flujo implícito**: la comunicación entre el complemento y el servicio en línea se implementa con código JavaScript del lado cliente.
- **Flujo de código de autorización**: la comunicación se establece *de servidor a servidor* entre la aplicación web del complemento y el servicio en línea. Por lo tanto, se implementa con código de servidor.

El propósito de los flujos es proteger la identidad y la autorización de la aplicación. En el flujo de código de autorización, se le proporciona un *secreto de cliente* que debe mantener oculto. Una aplicación de página única (SPA) no tiene manera de proteger el secreto, así que le recomendamos que use el flujo implícito para las SPA. 

Debería conocer las otras ventajas e inconvenientes de los dos flujos. Las definiciones oficiales, en [Código de autorización](https://tools.ietf.org/html/rfc6749#section-1.3.1) e [Implícito](https://tools.ietf.org/html/rfc6749#section-1.3.2), son un buen punto de partida. 

>**Nota:** También tiene la opción de hacer que un servicio intermediario se ocupe del proceso de autorización y pase el token de acceso al complemento. Para obtener más información, consulte la sección *Servicios intermediarios* más adelante en este artículo.

## Usar el flujo implícito en los complementos de Office
La mejor manera de saber si el servicio en línea admite el flujo implícito es consultar la documentación.

Para los servicios que lo admiten, proporcionamos una biblioteca de JavaScript que se encarga de todos los detalles de la tarea:

[Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

La carpeta \demo del repositorio contiene un complemento de ejemplo que usa la biblioteca para obtener acceso a varios servicios populares como Google, Facebook y Office 365.

Vea también la sección **Bibliotecas** más adelante en este artículo.

## Usar el flujo de código de autorización en los complementos de Office

Tenemos algunos complementos de muestra que usan el flujo de código de autorización:

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

Hay muchas bibliotecas disponibles para implementar el flujo de código de autorización en distintos lenguajes y marcos de trabajo. Para obtener más información, consulte la sección **Bibliotecas** más adelante en este artículo.

### Funciones de proxy o de transmisión

Puede usar el flujo de código de autorización incluso con una aplicación web sin servidor. Para ello, se almacenan los valores de *client ID* (id. de cliente) y *client secret* (secreto de cliente) en una función simple que se hospeda en un servicio como [Funciones de Azure](https://azure.microsoft.com/en-us/services/functions) o [Amazon Lambda](https://aws.amazon.com/lambda).
La función intercambia un código determinado por un *token de acceso* adecuado y lo transmite de vuelta al cliente. La seguridad de este método depende de lo protegido que esté el acceso a la función.

Para usar esta técnica, el complemento muestra una interfaz o una ventana emergente en la que se ve la pantalla de inicio de sesión del servicio en línea (Google, Facebook etc.). Cuando el usuario inicia sesión y concede al complemento permiso para acceder a sus recursos en el servicio en línea, el desarrollador recibe un código que se puede enviar a la función en línea. Los servicios descritos en la sección **Servicios intermediarios** de este artículo usan un flujo similar a este. 

## Bibliotecas

Hay bibliotecas disponibles para muchos lenguajes y plataformas y para ambos flujos. Algunas son de uso general y otras son para servicios en línea específicos. 

**Office 365 y otros servicios que usen Azure Active Directory como proveedor de autorización**: [bibliotecas de autenticación de Azure Active Directory](https://azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-libraries/). También está disponible una versión preliminar de la [Biblioteca de autenticación de Microsoft](https://www.nuget.org/packages/Microsoft.Identity.Client).

**Google**: busque "auth" o el nombre de su idioma en [GitHub.com/Google](https://github.com/google). La mayoría de los repositorios pertinentes se denominan `google-auth-library-[name of language]`.

**Facebook**: busque "biblioteca" o "sdk" en [Facebook for Developers](https://developers.facebook.com). 

**OAuth 2.0 en general**: hay una página de vínculos a bibliotecas para más de una docena idiomas gestionada por el grupo de trabajo de IETF para OAuth en: [OAuth Code](http://oauth.net/code/). Tenga en cuenta que algunas de estas bibliotecas son para implementar un servicio compatible con OAuth. Las bibliotecas de interés para un desarrollador de complementos se denominan bibliotecas *cliente*
en esta página porque el servidor web es un cliente del servicio compatible con OAuth.

## Servicios intermediarios

El complemento puede usar un servicio intermediario, como Auth0, que proporcione tokens de acceso para muchos servicios en línea populares o que simplifique el proceso de habilitar el inicio de sesión social para el complemento. Con muy poco código, el complemento puede usar script del lado cliente o código del lado servidor para conectarse al intermediario y enviar los tokens que necesita el servicio en línea. Todo el código que implementa la autorización se encuentra en el servicio intermediario. 

Tenemos un ejemplo que usa Auth0 para habilitar el inicio de sesión social con Facebook, Google y cuentas de Microsoft:

[Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)

## ¿Qué es CORS?

CORS es el acrónimo de [Cross Origin Resource Sharing](https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS) (Uso compartido de recursos entre orígenes). Para obtener información sobre cómo usar CORS dentro de los complementos, consulte [Addressing same-origin policy limitations in Office Add-ins](http://dev.office.com/docs/add-ins/develop/addressing-same-origin-policy-limitations).
