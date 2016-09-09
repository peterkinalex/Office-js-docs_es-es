
# Límites de recursos y optimización de rendimiento para los complementos de Office



Para crear la mejor experiencia para los usuarios, asegúrese de que su Office Add-in funcione dentro de límites específicos para el uso del núcleo de CPU y de la memoria, la confiabilidad y, en el caso de los complementos de Outlook, el tiempo de respuesta para evaluar expresiones regulares. Estos límites del uso de recursos de tiempo de ejecución se aplican a los complementos que se ejecutan en clientes de Office para Windows y OS X, pero no a Office Online, Outlook Web App ni OWA para dispositivos. También puede optimizar el rendimiento de sus complementos en dispositivos móviles y de escritorio optimizando el uso de recursos en el diseño y la implementación de los complementos.

## Límites de uso de recursos para complementos


Los límites de uso de recursos de tiempo de ejecución se aplican a todos los tipos de Office Add-ins. Estos límites ayudan a garantizar el rendimiento para sus usuarios y mitigar los ataques por denegación de servicio. Asegúrese de probar su Complemento de Office en la aplicación host de destino con una amplia gama de datos posibles y mida su rendimiento con los siguientes límites de uso de tiempo de ejecución:


-  **Uso de núcleos de CPU:** un único umbral de uso de núcleos de CPU del 90 %, que se comprueba tres veces en intervalos predeterminados de 5 segundos.
    
    El intervalo predeterminado para un cliente enriquecido de host para comprobar que el uso de núcleos de la CPU es cada 5 segundos. Si el cliente de host detecta que el uso de núcleos de la CPU de un complemento es superior al valor de umbral, muestra un mensaje preguntando si el usuario desea seguir ejecutando el complemento. Si el usuario decide continuar, el cliente de host no vuelve a preguntar al usuario durante esa sesión de edición. Es posible que los administradores deseen usar la clave de registro **AlertInterval** para aumentar el umbral y reducir la aparición de este mensaje de advertencia si los usuarios ejecutan complementos que consumen muchos recursos de la CPU.
    
-  **Uso de memoria:** umbral predeterminado de uso de memoria que se determina de forma dinámica según la memoria física disponible del dispositivo.
    
    De forma predeterminada, cuando un cliente enriquecido de host detecta que el uso de memoria física en un dispositivo supera el 80 % de la memoria disponible, el cliente empieza a supervisar el uso de la memoria del complemento, en un nivel de documento en búsqueda de complementos de contenido y panel de tareas, y en un nivel de buzón en búsqueda de complementos de Outlook. En un intervalo predeterminado de 5 segundos, el cliente advierte al usuario si el uso de memoria física para un conjunto de complementos en el nivel de documento o de buzón supera el 50 %. Este límite de uso de memoria usa memoria física en lugar de la memoria virtual para garantizar un rendimiento en dispositivos con memoria RAM limitada, como tabletas. Los administradores pueden reemplazar esta configuración dinámica con un límite explícito mediante la clave del registro de Windows **MemoryAlertThreshold** como una opción global, o ajustar el intervalo de alerta mediante la clave **AlertInterval** como un valor global.
    
-  **Tolerancia de bloqueos:** límite predeterminado de cuatro bloqueos por complemento.
    
    Los administradores pueden ajustar el umbral de bloqueos mediante el uso de la clave de registro **RestartManagerRetryLimit**.
    
-  **Bloqueo de aplicaciones:** umbral prolongado de falta de respuesta de 5 segundos por complemento.
    
    Esto afecta a la experiencia del usuario del complemento y la aplicación host. Cuando esto ocurre, la aplicación host reinicia automáticamente todos los complementos activos para un documento o un buzón (si procede), y advierte al usuario acerca de qué complemento dejó de responder. Los complementos pueden alcanzar este umbral cuando regularmente no realizan el procesamiento mientras se realizan tareas de larga duración. Existen técnicas para asegurarse de que no se produzca un bloqueo. Los administradores no pueden reemplazar este umbral.
    
     **Complementos de Outlook**
    
    Si algún complemento de Outlook supera los umbrales anteriores de uso del núcleo de CPU o de la memoria, o si excede el límite de tolerancia de bloqueos, Outlook deshabilita el complemento. El Centro de administración de Exchange muestra el estado deshabilitado de la aplicación.
    
     >**Nota** aunque solo los clientes enriquecidos de Outlook y no Outlook Web App u OWA para dispositivos supervisan el uso de recursos, si un cliente enriquecido deshabilita un complemento de Outlook, este complemento también se deshabilita para el uso en Outlook Web App y OWA para dispositivos.

    Además del núcleo CPU, la memoria y las reglas de confiabilidad, los complementos de Outlook deben respetar las siguientes reglas en la activación:
    
      -  **Tiempo de respuesta para las expresiones regulares:** Outlook tiene un umbral predeterminado de 1000 milisegundos para evaluar todas las expresiones regulares en el manifiesto de un complemento de Outlook. Si se sobrepasa ese umbral, Outlook volverá a intentar realizar la evaluación más adelante.
    
        Mediante una directiva de grupo o configuración específica de la aplicación en el registro de Windows, los administradores pueden ajustar este valor de umbral predeterminado de 1.000 milisegundos en la configuración **OutlookActivationAlertThreshold**. Para obtener más información, consulte [Reemplazar la configuración de uso de recursos para el rendimiento de los complementos de Office](http://msdn.microsoft.com/library/da14ec8c-5075-4035-a951-fc3c2b15c04b%28Office.15%29.aspx).
    
  -  **Reevaluación de expresiones regulares:** Outlook tiene un límite predeterminado de tres veces para volver a evaluar todas las expresiones regulares de un manifiesto. Si la evaluación produce errores las tres veces por superar el umbral establecido (que es el valor predeterminado de 1000 milisegundos o el valor establecido por **OutlookActivationAlertThreshold**, si esta configuración existe en el Registro de Windows), Outlook deshabilita el complemento de Outlook. En el Centro de administración de Exchange se muestra el estado deshabilitado y el uso del complemento en clientes avanzados de Outlook, en Outlook Web App y en OWA para dispositivos está deshabilitado.
    
    Mediante una directiva de grupo o configuración específica de la aplicación en el registro de Windows, los administradores pueden ajustar este número de veces para volver a intentar la evaluación en la configuración **OutlookActivationManagerRetryLimit**. Para obtener más información, consulte [Reemplazar la configuración de uso de recursos para el rendimiento de los complementos de Office](http://msdn.microsoft.com/library/da14ec8c-5075-4035-a951-fc3c2b15c04b%28Office.15%29.aspx).
    

    **Complementos de contenido y panel de tareas**
    
    Si algún complemento de panel de tareas o de contenido supera los umbrales anteriores de uso del núcleo de CPU y de la memoria, o si excede el límite de tolerancia de bloqueos, la aplicación host correspondiente muestra una advertencia al usuario. En este momento, el usuario puede hacer lo siguiente:
    
  - Reiniciar el complemento.
    
  - No aceptar más alertas de superación de ese umbral. Lo ideal sería que el usuario eliminara el complemento del documento, dado que continuar con el complemento podría provocar errores de estabilidad y rendimiento.
    

## Comprobación de los problemas de uso de recursos en el Registro de telemetría


Office proporciona un Registro de telemetría que mantiene un registro de determinados eventos (carga, apertura, cierre y errores) de las soluciones de Office que se ejecutan en el equipo local, incluidos los problemas de uso de recursos de una Complemento de Office. Si ha configurado el Registro de telemetría, puede usar Excel para abrir el Registro de telemetría en la siguiente ubicación predeterminada de su unidad local:

%Users%\ \<lt;Usuario actual \>gt; \AppData\Local\Microsoft\Office\15.0\Telemetry

Para cada evento de un complemento que sigue el Registro de telemetría, está la fecha y hora en la que ocurrió, el identificador del evento, la gravedad y un breve título descriptivo del evento, el nombre descriptivo y el identificador único del complemento y la aplicación que registró el evento. Puede actualizar el Registro de telemetría para ver los eventos que se siguen actualmente. La siguiente tabla contiene un par de ejemplos de complementos de Outlook que se siguieron en el Registro de telemetría. 



|**Fecha/Hora**|**Id. de evento**|**Severity**|**Título**|**Archivo**|**Id.**|**Aplicación**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|10/8/2012 5:57:10 p. m.|7||El manifiesto del complemento se descargó correctamente|Who's Who|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|10/8/2012 5:57:01 p. m.|7||El manifiesto del complemento se descargó correctamente|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|
 La siguiente tabla contiene los eventos de las Complementos de Office que sigue el Registro de telemetría en general.



|**Id. de evento**|**Título**|**Severity**|**Descripción**|
|:-----|:-----|:-----|:-----|
|7|El manifiesto del complemento se descargó correctamente||El manifiesto de la Complemento de Office se cargó correctamente y la aplicación host lo leyó correctamente.|
|8|No se pudo descargar el manifiesto del complemento|Crítico|La aplicación host no pudo cargar el archivo de manifiesto de la Complemento de Office desde el catálogo de SharePoint, el catálogo corporativo o la Tienda Office.|
|9|No se pudo analizar el marcado del complemento|Crítico|La aplicación host cargó el manifiesto de la Complemento de Office, pero no pudo leer el formato HTML de la aplicación.|
|10|El complemento usó demasiada CPU|Crítico|La Complemento de Office usó más del 90 % de los recursos de la CPU durante un período de tiempo limitado.|
|15|El complemento se deshabilitó porque se superó el tiempo de espera de búsqueda de cadenas||Los complementos de Outlook buscan en el mensaje y la línea de asunto de los correos electrónicos para averiguar si deben mostrarse con una expresión regular. Outlook deshabilitó el complemento de Outlook que aparece en la columna  **Archivo** porque agotó varias veces el tiempo de espera al intentar encontrar una expresión regular.|
|18|El complemento se cerró correctamente||La aplicación host pudo cerrar la Complemento de Office correctamente.|
|19|Error de tiempo de ejecución del complemento.|Crítico|Se produjo un problema en la Complemento de Office que provocó un error. Para obtener información detallada, vea el registro de  **Alertas de Microsoft Office** con el Visor de eventos de Windows en el equipo donde se produjo el error.|
|20|No se pudieron comprobar las licencias del complemento.|Crítico|La información de licencias de la Complemento de Office no se pudo comprobar y puede haber expirado. Para obtener información detallada, vea el registro de  **Alertas de Microsoft Office** con el Visor de eventos de Windows en el equipo donde se produjo el error.|
Para obtener más información, consulte [Implementación del Panel de telemetría](http://msdn.microsoft.com/en-us/library/f69cde72-689d-421f-99b8-c51676c77717%28Office.15%29.aspx) y [Solución de problemas de los archivos de Office y soluciones personalizadas con el registro de telemetría](http://msdn.microsoft.com/library/ef88e30e-7537-488e-bc72-8da29810f7aa%28Office.15%29.aspx)


## Técnicas de diseño e implementación


Si bien los límites de recursos para el uso de la CPU y la memoria, la tolerancia de bloqueos y la respuesta de la interfaz de usuario se aplican a las Complementos de Office cuando se ejecutan solo en clientes enriquecidos, la optimización del uso de estos recursos y de la batería debe ser una prioridad si desea que el complemento tenga un rendimiento satisfactorio en todos los clientes y dispositivos compatibles. La optimización es particularmente importante si el complemento realiza operaciones de ejecución prolongada o administra grandes conjuntos de datos. En la siguiente lista se sugieren algunas técnicas para dividir las operaciones con uso intensivo de datos o CPU en fragmentos más pequeños para que el complemento pueda evitar el consumo excesivo de recursos y la aplicación host pueda seguir respondiendo con normalidad:


- En un escenario en el que el complemento necesita leer un gran volumen de datos de un conjunto de datos sin enlazar, puede aplicar la paginación cuando se leen los datos de una tabla o reducir el tamaño de los datos en operaciones de lectura más breves, en lugar de intentar completar la lectura en una sola operación. 
    
    Para un ejemplo de código JavaScript y jQuery que muestra la interrupción de una serie de operaciones de entrada y salida potencialmente prolongada y que consume muchos recursos de CPU, consulte [¿Cómo puedo dar volver a dar control (brevemente) al explorador durante procesamiento intensivo de JavaScript?](http://stackoverflow.com/questions/210821/how-can-i-give-control-back-briefly-to-the-browser-during-intensive-javascript). Este ejemplo usa el método [setTimeout](http://msdn.microsoft.com/en-us/library/ie/ms536753%28v=vs.85%29.aspx) del objeto global para limitar la duración de entrada y salida. También controla los datos en fragmentos definidos en lugar de datos sin delimitar de forma aleatoria.
    
- Si el complemento usa un algoritmo con uso intensivo de CPU para procesar un gran volumen de datos, puede usar Web Workers para realizar la tarea de ejecución prolongada en segundo plano mientras se ejecuta un script individual en primer plano, como mostrar el progreso en la interfaz de usuario. Los Web Workers no bloquean las actividades del usuario y permiten que la página HTML siga respondiendo con normalidad. Para ver un ejemplo de Web Workers, vea [Introducción a los Web Workers](http://www.mdl5rocks.com/en/tutorials/workers/basics/). Vea [Web Workers](http://msdn.microsoft.com/en-us/library/IE/hh772807%28v=vs.85%29.aspx) para obtener más información sobre la API de Web Workers de Internet Explorer.
    
- Si el complemento usa un algoritmo con uso intensivo de CPU, pero puede dividir la entrada o salida de los datos en conjuntos más pequeños, considere crear un servicio web, transferir los datos al servicio web para quitar la carga de la CPU y esperar una devolución de llamada asincrónica.
    
- Pruebe el complemento con el mayor volumen de datos esperado y restrinja el procesamiento que puede realizar el complemento hasta ese límite.
    

## Recursos adicionales



- [Privacidad y seguridad de complementos para Office](../../docs/develop/privacy-and-security.md)
    
- [Límites para la activación y API de JavaScript para complementos de Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
