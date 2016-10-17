
# <a name="authenticate-an-outlook-add-in-by-using-exchange-identity-tokens"></a>Autenticar un complemento de Outlook con los tokens de identidad de Exchange

Su complemento de Outlook puede proporcionar a sus clientes información de cualquier parte de Internet, ya sea desde el servidor que hospeda el complemento, desde su red interna o desde alguna otra parte de la nube. Si esa información está protegida, su complemento de correo necesitará una forma de asociar la cuenta de correo de Exchange con su servicio de información. Exchange 2013 puede habilitar el sistema de inicio de sesión único (SSO) para el complemento proporcionando un token que identifique la cuenta de correo que hace la solicitud. Puede asociar este token con un usuario registrado del complemento de manera que se pueda reconocer al usuario siempre que el complemento de Outlook se conecte al servicio.

## <a name="identity-tokens"></a>Tokens de identidad


Dos de los complementos de ejemplo usan información que está disponible al público. Uno de ellos muestra un mapa de Bing para las direcciones de un mensaje y la otra muestra una vista previa de los vínculos de vídeo de YouTube de un mensaje. Pero su complemento de Outlook también puede obtener acceso a información que no sea de acceso público. Puede usar el servidor que hospeda el complemento para vincularlo con la información de la red interna o de cualquier parte de la nube.

Puede usar muchas técnicas distintas para identificar y autenticar los usuarios del complemento. Exchange 2013 simplifica la autenticación de los usuarios proporcionando al complemento de Outlook un token de identidad que identifica a una cuenta de correo de Exchange específica. Puede asociar el token en el servicio con un usuario registrado, lo que permite el inicio de sesión único (SSO) para los clientes que usan complementos de Outlook. 

Para usar SSO en el complemento, el código hace lo siguiente:


* Llama a una función en la API del complemento de Outlook que devuelve un token de identidad.
* Envía el token a su servidor junto con una solicitud.
* Abre la respuesta recibida desde el servidor para mostrar la información del servicio.
    
Del lado servidor, la cosa se complica un poco. Cuando el servidor recibe una solicitud de un complemento de Outlook, el proceso es así:

* El servidor valida el token. Puede usar la [biblioteca de validación de tokens administrados](../../docs/outlook/use-the-token-validation-library.md), o bien puede [crear su propia biblioteca](../../docs/outlook/validate-an-identity-token.md) para el servicio.
* El servidor busca el identificador único del token para ver si está asociado a alguna identidad conocida. El servicio tiene que [implementar un método que haga coincidir el identificador](../../docs/outlook/authenticate-a-user-with-an-identity-token.md) con usuarios conocidos del servicio.
* Si el identificador único coincide con un identificador previamente almacenado con un conjunto de credenciales en el servidor, el servidor puede responder con la información solicitada sin pedirle al cliente que inicie sesión en el servicio.
* Si el identificador único es desconocido, el servidor envía una respuesta solicitándole al usuario que inicie la sesión con credenciales para el servidor.
* Si las credenciales coinciden con alguna identidad conocida en el servidor, puede asociar esa identidad al identificador único en el token de modo que la próxima vez que se reciba una solicitud, el servidor pueda responder sin solicitar un paso adicional de inicio de sesión.

 >**Nota**  Esto es una sugerencia sobre cómo usar el token de identidad. Como siempre, cuando esté trabajando con identidades y autenticaciones, deberá asegurarse de que el código cumpla los requisitos de seguridad de su organización.

Veamos los detalles. En los artículos siguientes, usaremos un sencillo complemento de Outlook que envíe a un servicio web el token de identidad y una lista de números de teléfono que se encuentran en el mensaje. 

- [Contenido del token de identidad de Exchange](../outlook/inside-the-identity-token.md)
- [Llamar a un servicio de un complemento de Outlook con un token de identidad en Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
- [Usar la biblioteca de validación de tokens de Exchange](../outlvalidate-an-identity-token.md ook/use-the-token-validation-library.md)
- [Validar un token de identidad de Exchange](../outlook/validate-an-identity-token.md )
- [Autenticar un usuario con un token de identidad para Exchange](../outlook/validate-an-identity-token.md)


## <a name="additional-resources"></a>Recursos adicionales



- [Complementos de Outlook](../outlook/outlook-add-ins.md)
    
- [Llamar a servicios web desde un complemento de Outlook](../outlook/web-services.md)
    


