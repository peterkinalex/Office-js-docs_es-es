
# Límites de activación y API de JavaScript para complementos de Outlook

Para proporcionar una experiencia satisfactoria a los usuarios de complementos de Outlook, debe tener en cuenta ciertas recomendaciones de activación y uso de la API, e implementar los complementos dentro de estos límites establecidos. Estas recomendaciones existen para que un complemento individual no necesite que Exchange Server o Outlook pasen períodos de tiempo inusualmente largos procesando sus reglas de activación o llamadas a la API de JavaScript para Office. Eso afectaría a la experiencia global de usuario para Outlook y otros complementos. Los límites se aplican al diseño de reglas de activación en el manifiesto del complemento y el uso de propiedades personalizadas, la configuración de movilidad, los destinatarios, las solicitudes y respuestas de Servicios Web Exchange (EWS) y las llamadas asincrónicas. 

 >**Nota** Si el complemento se ejecuta en un cliente avanzado de Outlook, tiene que comprobar también que el complemento se ejecuta dentro de ciertos límites de uso de recursos de tiempo de ejecución. 


## Límites para las reglas de activación


Siga las instrucciones que se especifican a continuación al diseñar reglas de activación para complementos de Outlook:


- Limite el tamaño del manifiesto a 256 KB. Si excede ese límite, no podrá instalar el complemento de Outlook para un buzón de Exchange.

- Especifique hasta 15 reglas de activación para el complemento. Si excede ese límite, no podrá instalar el complemento.
    
- Si usa una regla [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) en el cuerpo del elemento seleccionado, puede esperar que un cliente enriquecido de Outlook aplique la regla solo en el primer MB del cuerpo y no para el resto del cuerpo si se supera ese límite. El complemento no se activaría si encuentra una coincidencia solo después del primer MB del cuerpo. Si espera que ese sea un escenario probable, vuelva a diseñar las condiciones para la activación.
    
- Si usa expresiones regulares en las reglas **ItemHasKnownEntity** y [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx), tenga en cuenta que los límites y las directrices siguientes que se suelen aplicar a cualquier host de Outlook y los que se describen en las tablas 1, 2 y 3 diferirán según el host:
    
      - Especifique hasta cinco expresiones regulares en las reglas de activación de un complemento. Si supera ese límite, no podrá instalar ningún complemento.
    
  - Especifique expresiones regulares para que los resultados que anticipe se encuentren entre los 50 primeros con la llamada de método **getRegExMatches**.
    
  - Puede especificar aserciones look-ahead en expresiones regulares, pero no look-behind (?<=text) ni look-behind negativo (?<!text).
    

En la tabla 1 se indican los límites y se describen las diferencias de compatibilidad de las expresiones regulares existentes entre un cliente avanzado de Outlook y Outlook Web App u OWA para dispositivos. Esta compatibilidad es independiente del tipo de dispositivo y del cuerpo del elemento.


 **The support is independent of any specific type of device and item body.**


|**General differences in the support for regular expressions**|**Cliente enriquecido de Outlook**|
|:-----|:-----|
|Usa el motor de expresiones regulares C++ proporcionado como parte de la biblioteca de plantillas estándar de Visual Studio. Este motor cumple con los estándares de ECMAScript 5. |Usa la evaluación de expresiones regulares que forma parte de JavaScript, la proporciona el explorador y admite un superconjunto de ECMAScript 5.|
|Debido a los motores de regex diferentes, debe esperar que un regex que incluye una clase de carácter personalizado basada en clases de caracteres predefinidas pueda devolver resultados diferentes en un cliente enriquecido de Outlook que en Outlook Web App u OWA para dispositivos.<br/><br/>Por ejemplo, el regex "[\s\S]{0,100}" coincide con cualquier número, entre 0 y 100, de caracteres únicos que sea un espacio en blanco o que no sea un espacio en blanco. Este regex devuelve resultados diferentes en un cliente enriquecido de Outlook que en Outlook Web App y OWA para dispositivos. Debería reescribir el regex como ""(\s\|\S){0,100}" como una solución temporal. Este regex de solución temporal coincide con cualquier número, entre 0 y 100, de espacios en blanco o que no sean espacios en blanco.<br/><br/>Pruebe cada expresión regular detenidamente en cada host de Outlook y, si una de ellas devuelve resultados diferentes, reescríbala. |Pruebe cada expresión regular detenidamente en cada host de Outlook y, si una de ellas devuelve resultados diferentes, reescríbala.|
|De manera predeterminada, limita la evaluación de todas las expresiones regulares de un complemento en un segundo. Si se excede este límite, la evaluación se vuelve a iniciar hasta tres veces. Cuando se supera este segundo límite, un cliente enriquecido de Outlook deshabilita el complemento para que no se ejecute para el mismo buzón en ninguno de los hosts de Outlook.<br/><br/>Los administradores pueden reemplazar estos límites de evaluación mediante las claves del registro **OutlookActivationAlertThreshold** y **OutlookActivationManagerRetryLimit**.|No admiten la misma supervisión de recursos ni la misma configuración del Registro que un cliente enriquecido de Outlook. Pero los complementos con expresiones regulares que necesitan un tiempo de evaluación excesivo en un cliente enriquecido de Outlook se desactivan para el mismo buzón en todos los hosts de Outlook.|

En la tabla 2 se indican los límites y se describen las diferencias en la parte del cuerpo de elemento donde Outlook aplica una expresión regular. Algunos de estos límites dependen del tipo de dispositivo y del cuerpo del elemento, si la expresión regular se aplica en el cuerpo del elemento.

**Tabla 2. Límites de tamaño del cuerpo de elemento evaluado**


||**Cliente enriquecido de Outlook**|**Cliente enriquecido de Outlook**|**Outlook Web App**|
|:-----|:-----|:-----|:-----|
|Factor de forma|Cualquier dispositivo compatible|Smartphones Android, iPad o iPhone|Cualquier dispositivo compatible que no sea un smartphone Android, iPad y iPhone|
|Cuerpo de elemento en texto sin formato|Aplica el regex al primer 1 MB de los datos del cuerpo, pero no al resto del cuerpo por encima de este límite.|Activa el complemento solo si el cuerpo es inferior a 16.000 caracteres.|Activa el complemento solo si el cuerpo es inferior a 500 000 caracteres.|
|Cuerpo de elemento en HTML|Aplica el regex a los primeros 512 KB de los datos del cuerpo, pero no al resto del cuerpo por encima de este límite. El número real de caracteres depende del cifrado, que puede oscilar entre 1 y 4 bytes por carácter.|Aplica el regex a los primeros 64 000 caracteres (caracteres de etiqueta HTM inclusive), pero no al resto del cuerpo por encima de este límite.|Activa el complemento solo si el cuerpo es inferior a 500 000 caracteres.|

En la tabla 3 se indican los límites y se describen las diferencias de los resultados que los hosts de Outlook devuelven después de evaluar una expresión regular. La compatibilidad es independiente del tipo de dispositivo, pero puede depender del tipo de cuerpo de elemento si la expresión regular se aplica en dicho cuerpo de elemento.

**Tabla 3. Límites de los resultados devueltos**


||**Cliente enriquecido de Outlook**|**Cliente enriquecido de Outlook**|
|:-----|:-----|:-----|
|Orden de las coincidencias devueltas|Presupone que los resultados que devuelve  **getRegExMatches** para la misma expresión regular aplicada al mismo elemento son distintos en un cliente enriquecido de Outlook, por un lado, y en Outlook Web App o bien OWA para dispositivos, por otro.|Presupone que los resultados que devuelve  **getRegExMatches** para la misma expresión regular aplicada al mismo elemento son distintos en un cliente enriquecido de Outlook, por un lado, y en Outlook Web App o bien OWA para dispositivos, por otro.|
|Cuerpo de elemento en texto sin formato|**getRegExMatches** devuelve hasta 50 resultados que tengan un máximo de 1.536 caracteres (1,5 KB).<br/><br/>**Nota**: **getRegExMatches** no devuelve resultados en un orden específico en la matriz devuelta. En general, presuponga que el orden de coincidencias en un cliente enriquecido de Outlook para la misma expresión regular aplicada al mismo elemento son distintos de los de Outlook Web App y OWA para dispositivos.|**getRegExMatches** devuelve hasta 50 resultados que tengan un máximo de 3.072 caracteres (3 KB).|
|Cuerpo de elemento en HTML|**getRegExMatches** devuelve hasta 50 resultados que tengan un máximo de 3.072 caracteres (3 KB).<br/> <br/> **Nota**: **getRegExMatches** no devuelve resultados en un orden específico en la matriz devuelta. En general, presuponga que el orden de coincidencias en un cliente enriquecido de Outlook para la misma expresión regular aplicada al mismo elemento son distintos de los de Outlook Web App y OWA para dispositivos.|**getRegExMatches** devuelve hasta 50 resultados que tengan un máximo de 3.072 caracteres (3 KB).|

## Límites para la API de JavaScript


Además de las directrices anteriores para las reglas de activación, cada uno de los hosts Outlook impone ciertos límites en el modelo de objeto de JavaScript, como se describe en la Tabla 4.


**Además de las directrices anteriores para las reglas de activación, cada uno de los hosts Outlook impone ciertos límites en el modelo de objeto de JavaScript, como se describe en la Tabla 4.**


|**Característica**|**Límite**|**Límite**|**Descripción**|
|:-----|:-----|:-----|:-----|
|Propiedades personalizadas|2500 caracteres|Objeto [CustomProperties](../../reference/outlook/CustomProperties.md)<br/> <br/>Método [item.loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md)|Límite para todas las propiedades personalizadas de un elemento de cita o mensaje. Todos los hosts Outlook devuelven un error si el tamaño total de todas las propiedades personalizadas de un complemento supera este límite.|
|Configuración de roaming|32 KB de caracteres|Objeto [RoamingSettings](../../reference/outlook/RoamingSettings.md).<br/><br/> Propiedad [context.roamingSettings](../../reference/outlook/Office.context.md)|Límite de todas las configuraciones de itinerancia para el complemento. Todos los hosts Outlook devuelven un error si su configuración supera este límite.|
|Extraer entidades conocidas|2000 caracteres|Método [item.getEntities](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Método [item.getEntitiesByType](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Método [item.getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md)|Límite de Exchange Server para extraer las entidades conocidas del cuerpo de elemento. Exchange Server pasa por alto las entidades una vez superado ese límite. Tenga en cuenta que el límite es independiente de si el complemento usa una regla  **ItemHasKnownEntity**.|
|Servicios web de Exchange|1 MB de caracteres|Método [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)|Método **mailbox.makeEwsRequestAsync**|
|Destinatarios|100 destinatarios|Propiedad [item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Propiedad [item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Propiedad [item.resources](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Propiedad [item.to](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Propiedad [item.cc](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Método [Recipients.addAsync](../../reference/outlook/Recipients.md)<br/> <br/>Método [Recipient.getAsync](../../reference/outlook/Recipients.md)<br/> <br/>Método [Recipient.setAsync](../../reference/outlook/Recipients.md)|Límite de los destinatarios especificados en cada propiedad.|
|Nombre para mostrar|255 caracteres|Propiedad [EmailAddressDetails.displayName](../../reference/outlook/simple-types.md)<br/><br/> Objeto [Recipients](../../reference/outlook/Recipients.md)<br/><br/> Propiedad **item.requiredAttendees**<br/><br/> Propiedad **item.optionalAttendees** <br/><br/>Propiedad **item.resources** <br/><br/>Propiedad **item.to** <br/><br/>Propiedad **item.cc**|Límite de longitud de un nombre para mostrar en una cita o mensaje.|
|Definir el asunto|255 caracteres|Método [mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)<br/><br/> Método [Subject.setAsync](../../reference/outlook/Subject.md)|Límite del asunto en el nuevo formulario de cita o para definir el asunto de una cita o mensaje.|
|Definir la ubicación|255 caracteres|Método [Location.setAsync](../../reference/outlook/Location.md)|Límite para definir la ubicación de una cita o convocatoria de reunión|
|Cuerpo de un nuevo formulario de cita|32 KB de caracteres|Método **Mailbox.displayNewAppointmentForm**|Límite del cuerpo de un nuevo formulario de cita.|
|Mostrar el cuerpo de un elemento existente|32 KB de caracteres|Método [mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md)<br/><br/> Método [mailbox.displayMessageForm](../../reference/outlook/Office.context.mailbox.md)|Para Outlook Web App y OWA para dispositivos: límite del cuerpo del formulario de cita o mensaje existente.|
|Definir el cuerpo|1 MB de caracteres|Método [Body.prependAsync](../../reference/outlook/Body.md)<br/> <br/>[Body.setAsync](../../reference/outlook/Body.md)<br/><br/>Método [Body.setSelectedDataAsync](../../reference/outlook/Body.md)|Límite para definir el cuerpo de un elemento de mensaje o cita.|
|Número de archivos adjuntos|499 archivos en Outlook Web App y OWA para dispositivos|Método [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)|Método **item.addFileAttachmentAsync**|
|Tamaño de datos adjuntos|Depende de Exchange Server|Método **item.addFileAttachmentAsync**|Existe un límite para el tamaño de todos los datos adjuntos de un elemento que un administrador puede configurar en el Exchange Server del buzón de correo del usuario. Para un cliente avanzado de Outlook, esto limita el número de datos adjuntos por elemento. Para Outlook Web App y OWA para dispositivos, el menor de los dos límites (el número de datos adjuntos y el tamaño de los datos adjuntos) define la restricción de los datos adjuntos de un elemento.|
|Nombres de archivos adjuntos|255 caracteres|Método **item.addFileAttachmentAsync**|Límite de longitud del nombre de los archivos adjuntos que se van a agregar a un elemento.|
|URI de datos adjuntos|2048 caracteres|Método **item.addFileAttachmentAsync**|Límite del URI del nombre de archivo que se va a agregar como datos adjuntos a un elemento.|
|Identificador de datos adjuntos|100 caracteres|Método [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)<br/><br/> Método [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)|Límite de longitud del identificador de los datos adjuntos que se van a agregar o eliminar de un elemento.|
|Llamadas asincrónicas|3 llamadas|Método **item.addFileAttachmentAsync**<br/><br/>Método **item.addItemAttachmentAsync**<br/><br/><br/>Método **item.removeAttachmentAsync**<br/><br/> Método [Body.getTypeAsync](../../reference/outlook/Body.md)<br/><br/>Método **Body.prependAsync**<br/><br/>Método **Body.setSelectedDataAsync**<br/><br/> Método [CustomProperties.saveAsync](../../reference/outlook/CustomProperties.md)<br/><br/><br/> Método [item.LoadCustomPropertiesAysnc](../../reference/outlook/Office.context.mailbox.item.md)<br/><br/><br/> Método [Location.getAsync](../../reference/outlook/Location.md)<br/><br/>Método **Location.setAsync**<br/><br/> Método [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md)<br/><br/> Método [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md)<br/><br/> Método [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)<br/><br/>Método **Recipients.addAsync**<br/><br/> Método [Recipients.getAsync](../../reference/outlook/Recipients.md)<br/><br/>Método **Recipients.setAsync**<br/><br/> Método [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md)<br/><br/> Método [Subject.getAsync](../../reference/outlook/Subject.md)<br/><br/>Método **Subject.setAsync**<br/><br/> Método [Time.getAsync](../../reference/outlook/Time.md)<br/><br/> Método [Time.setAsync](../../reference/outlook/Time.md)|Para Outlook Web App o OWA para dispositivos: límite del número de llamadas asincrónicas por vez, ya que los exploradores solo permiten un número limitado de llamadas asincrónicas a los servidores. |

## Recursos adicionales



- [Implementar e instalar complementos de Outlook para probarlos](../outlook/testing-and-tips.md)
    
- [Privacidad, permisos y seguridad para los complementos de Outlook](../outlook/../../docs/develop/privacy-and-security.md)
    
