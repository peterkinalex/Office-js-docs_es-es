# Modelos de diseño de la experiencia del usuario para complementos de Office. 

Al diseñar complementos de Office, el diseño de la experiencia del usuario del complemento tiene que proporcionar experiencias atractivas que amplíen las funciones de Office. Para crear un gran complemento, este tiene que proporcionar una experiencia de primera ejecución, una experiencia del usuario de primera clase y transiciones suaves entre las páginas, entre otras cosas. Proporcionar una experiencia del usuario moderna y sin complicaciones aumenta la retención de usuarios y la adopción del complemento. En este artículo se presentan recursos de experiencia del usuario para diseñadores y desarrolladores donde:

* Se describen modelos de diseño de la experiencia del usuario comunes basados en procedimientos recomendados.
* Se implementan componentes y estilos de Office Fabric.
* Se implementan complementos que parecen una extensión natural de la interfaz de usuario predeterminada de Office. 

## ¿Cómo empiezo a usar los recursos de ejemplo de diseño de complementos de Office?

No existen requisitos previos para usar estos recursos de diseño o código. Para empezar a crear una excelente experiencia del usuario para su complemento:

* Revise los modelos de diseño de la experiencia del usuario e identifique los que son importantes para su complemento. Por ejemplo, seleccione una de las experiencias de primera ejecución.
* Después, siga uno o más de estos procedimientos:
	* Copie los archivos de código en el proyecto del complemento y empiece a personalizarlos para adaptarlos a sus requisitos. Necesitará el [archivo common.js](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/), la [carpeta de recursos](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets) y la carpeta de código para el modelo de diseño que necesite. Vea los vínculos a continuación.
	* Descargue los archivos PDF de referencia y úselos como guía al crear su propio diseño de la experiencia del usuario. Vea los vínculos a continuación.
	* Descargue los archivos de Adobe Illustrator y modifíquelos para crear un boceto de sus propios diseños de complemento. Descárguelos desde [aquí](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Source%20Files).
 

## Primera ejecución

Una experiencia de primera ejecución es la experiencia que tiene un usuario al abrir el complemento por primera vez. En la lista siguiente se muestran los modelos de diseño que puede incluir en el complemento. A continuación se muestran imágenes de todos los modelos.

* 
            En **Primeros pasos** se proporciona a los usuarios una lista ordenada de pasos para empezar a usar su complemento. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_StepsToStart.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/instruction-step))
* 
            En **Valor** se comunica la propuesta de valor del complemento. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_ValuePlacemat.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/value-placemat))
* 
            En **Vídeo** se muestra a los usuarios un vídeo antes de que empiecen a usar el complemento. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_VideoPlacemat.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/video-placemat))
* 
            En el **tutorial**, se muestra a los usuarios una serie de características o información antes de que empiecen a usar el complemento. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_PagingPanel.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/walkthrough))
* La [Tienda Office](https://msdn.microsoft.com/es-es/library/office/jj220033.aspx) tiene un sistema para proporcionar a los usuarios una versión de prueba de un complemento pero, si quiere tener control total sobre la interfaz de usuario para una experiencia de prueba, use las plantillas siguientes:
	* 
            En **Prueba** se muestra a los usuarios cómo empezar a usar una versión de prueba del complemento. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_TrialVersion.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat))
	* 
            En **Característica de prueba** se informa a los usuarios de que la característica que intentan usar no está disponible en la versión de prueba del complemento. ([código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat-feature))


> Nota: Determine si es importante para su escenario mostrar una o varias veces la experiencia de primera ejecución a los usuarios. Por ejemplo, si los usuarios usan el complemento de vez en cuando, es posible que se olviden de cómo usarlo. Volver a ver la experiencia de primera ejecución puede resultar útil para esos usuarios. 

 <table>
 <tr><th>Pasos para empezar</th><th>Valor</th><th>Vídeo</th></tr>
 <tr><td>![instruction steps" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/instruction.step.PNG)</td><td>![value placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/value.placemat.PNG)</td><td>![video placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/video.placemat.PNG)</td></tr>
 </table>

 <table>
 <tr><th>Primera página del tutorial</th><th>Prueba</th><th>Característica de prueba</th></tr>
 <tr><td>![walkthrough 1" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/walkthrough1.PNG)</td><td>![trial placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.PNG)</td><td>![trial placemat feature" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.feature.PNG)</td></tr>
 </table> 


## Genérico y con personalización de marca

* **Página de aterrizaje** es la primera página que visitan los usuarios después de la experiencia de primera ejecución o después de un proceso de inicio de sesión. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Helpful%20Templates/AddIn_Template_Standard_Layout.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/generic/landing-page))

<table>
 <tr><th>Aterrizaje</th></tr>
 <tr><td>![landing page" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/landing.page.PNG)</td></tr>
 </table>

## Notificaciones

Hay varias formas en que el complemento puede notificar a los usuarios de eventos, como errores o progreso. En la lista siguiente se muestran estas técnicas. A continuación se muestran imágenes de todos los modelos.

* **Diálogo insertado**: muestra un diálogo dentro del panel de tareas con información y, de manera opcional, una experiencia interactiva con botones u otros controles. Se puede usar para mostrar una notificación al usuario o para confirmar una acción. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Embedded_Dialog.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/embedded-dialog))
* **Mensaje en línea**: indica un error, una operación correcta o información, y puede aparecer en una ubicación especificada del panel de tareas. Por ejemplo, si un usuario escribe en un cuadro de texto una dirección de correo electrónico con un formato incorrecto, se muestra un mensaje de error justo debajo del cuadro de texto. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_Inline_Message.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/inline-message))
* **Mensaje emergente**: proporciona información y, de manera opcional, una sencilla llamada a la acción en un mensaje emergente que se puede contraer a una única línea, expandir a varias líneas o descartar. Puede usar mensajes emergentes para informar sobre una actualización de servicio o para mostrar una sugerencia útil cuando se inicie el complemento. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_messagebanner.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/message-banner))
* **Barra de progreso**: indica el progreso de un proceso sincrónico de ejecución prolongada, como una tarea de configuración que es necesario completar antes de que el usuario pueda realizar otra acción. Es una página intersticial independiente que también refuerza la marca del complemento. Use una barra de progreso cuando el proceso pueda enviar medidas periódicas del progreso para informar al complemento. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/progress-bar))
* **Indicador giratorio**: indica que se está realizando un proceso sincrónico de ejecución prolongada, pero no proporciona ninguna indicación de cuánto se ha completado. Es una página intersticial independiente que también refuerza la marca del complemento. Use un indicador giratorio cuando el complemento no pueda determinar de forma confiable qué parte del proceso se ha completado. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/spinner))
* **Notificación del sistema**: proporciona un mensaje breve que desaparece después de unos segundos. Como es posible que el usuario no vea el mensaje, use la notificación del sistema solo para información que no se considere esencial. Es una opción adecuada para notificar a los usuarios de un evento en un sistema remoto, como el recibo de un correo electrónico. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_toast.pdf "PDF"), [código](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/toast))

 <table>
 <tr><th>Diálogo insertado</th><th>Mensaje en línea</th><th>Mensaje emergente</th></tr>
 <tr><td>![embedded dialog" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/embedded.dialog.PNG)</td><td>![inline message" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/inline.message.PNG)</td><td>![message banner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/message.banner.PNG)</td></tr>
 </table>

 <table>
 <tr><th>Barra de progreso</th><th>Indicador giratorio</th><th>Notificación del sistema</th></tr>
 <tr><td>![progress bar" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/progress.bar.PNG)</td><td>![spinner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/spinner.PNG)</td><td>![toast" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/toast.PNG)</td></tr>
 </table>

## Problemas conocidos

* Al ejecutar algunos archivos de código fuera de un proyecto de complemento, se muestra un error de JavaScript. 
	* Solución: Asegúrese de agregar los archivos a un proyecto de complemento de Office. 
	
## Recursos adicionales

* [Procedimientos recomendados para desarrollar complementos de Office](https://dev.office.com/docs/add-ins/design/add-in-development-best-practices)
* [Office UI Fabric](http://dev.office.com/fabric/)
