
# <a name="requirements-for-running-office-add-ins"></a>Requisitos para ejecutar complementos de Office


En este artículo se describen los requisitos de software y de dispositivo para ejecutar complementos de Office.

>**Nota:** Para consultar una lista detallada de dónde se pueden usar actualmente los complementos de Office, vea la página [Disponibilidad de plataformas y hosts de los complementos de Office](http://dev.office.com/add-in-availability). 


## <a name="server-requirements"></a>Requisitos de servidor

Para poder instalar y ejecutar cualquier Complemento de Office., primero debe implementar los archivos de manifiesto y página web para la interfaz de usuario y el código de su complemento en las ubicaciones de servidor correspondientes.

Para todos los tipos de complementos (contenido, Outlook, complementos de panel de tareas y comandos de complemento), tendrá que implementar los archivos de página web del complemento en un servidor web o un servicio de hospedaje de sitios web, como [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).


 >**Nota:** Cuando desarrolla y depura un complemento en Visual Studio, Visual Studio implementa y ejecuta los archivos de la página web de su aplicación de forma local con IIS Express y no necesita un servidor web adicional. De forma similar, cuando desarrolla y depura con Napa en el explorador, implementa y ejecuta los archivos de página web de su aplicación desde el almacenamiento asociado con la cuenta usada para iniciar sesión en Napa.

Para complementos de panel de tareas y contenido, en las aplicaciones host de Office (aplicaciones web de Access, Word, Excel, PowerPoint o Project), también necesita un [catálogo de complementos](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) en SharePoint para cargar el archivo de manifiesto XML del complemento.

Para probar y ejecutar un complemento de Outlook, la cuenta de correo electrónico de Outlook del usuario debe encontrarse en Exchange 2013 o una versión posterior, disponible a través de Office 365, Exchange Online, o una instalación local. El usuario o el administrador instalarán los archivos de manifiesto para los complementos de Outlook en dicho servidor.

 >**Nota:**   Las cuentas de correo electrónico POP e IMAP de Outlook no admiten Complementos de Office.




## <a name="client-requirements:-windows-desktop-and-tablet"></a>Requisitos de cliente: dispositivo de escritorio y tableta con Windows

Para desarrollar un Complemento de Office, se necesita el siguiente software en los clientes de escritorio de Office o clientes web que ejecuten dispositivos de escritorio, portátiles o tabletas basados en Windows:


- Para equipos de escritorio x86 y x64 con Windows y tabletas como Surface Pro:

    - La versión de 32 bits o de 64 bits de Office 2013 o una versión posterior, en ejecución en Windows 7 o una versión posterior.

    - Excel 2013, Outlook 2013, PowerPoint 2013, Project Profesional 2013, Project 2013 SP1, Word 2013 o una versión posterior del cliente de Office, si desea probar o ejecutar un Complemento de Office específicamente para uno de estos clientes de escritorio de Office. Los clientes de escritorio de Office pueden instalarse de forma local o a través de Hacer clic y ejecutar en el equipo cliente.

- Internet Explorer 9 o posterior (se debe instalar, pero no es necesario que sea el explorador predeterminado). Para admitir Complementos de Office, el cliente de Office que actúa como host usa componentes del explorador que forman parte de Internet Explorer 9 o de una versión posterior.

- Uno de los siguientes exploradores predeterminados: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13 o una versión posterior de cualquiera de ellos.

- Un editor de HTML y JavaScript, como el Bloc de notas, [Visual Studio y Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs) o una herramienta de desarrollo web de terceros.


## <a name="client-requirements:-os-x-desktop"></a>Requisitos de cliente: escritorio de OS X

Outlook para Mac, que se distribuye como parte de Office 365, admite complementos de Outlook. La ejecución de complementos de Outlook en Outlook para Mac tiene los mismos requisitos que el propio Outlook para Mac: el sistema operativo debe ser como mínimo OS X v10.10 "Yosemite". Ya que Outlook para Mac usa WebKit como motor de diseño para presentar las páginas del complemento, no existen dependencias de explorador adicionales.

Estas son las versiones mínimas del cliente de Office para Mac que admiten complementos de Office:
- Word para Mac versión 15.18 (160109) 
- Excel para Mac versión 15.19 (160206) 
- PowerPoint para Mac versión 15.24 (160614)

## <a name="client-requirements:-browser-support-for-office-online-web-clients-and-sharepoint"></a>Requisitos de cliente: compatibilidad del explorador con clientes web de Office Online y SharePoint

Cualquier explorador que admita ECMAScript 5.1, HTML5 y CSS3, como Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6 o una versión posterior de estos exploradores.


## <a name="client-requirements:-non-windows-smartphone-and-tablet"></a>Requisitos de cliente: smartphone y tabletas sin Windows

Se necesita el siguiente software para probar y ejecutar complementos de Outlook específicamente para OWA para dispositivos y Outlook Web App cuando se ejecutan en un explorador en smartphones o tabletas que no son Windows.


| Aplicación host | Dispositivo | Sistema operativo | Cuenta de Exchange | Explorador móvil |
|:-----|:-----|:-----|:-----|:-----|
|OWA para Android|Smartphones Android. Técnicamente, [Android OS](https://developer.android.com/guide/practices/screens_support.html) considera estos dispositivos "pequeños" o "normales".|Android 4.4 KitKat o posterior|En la última actualización de Office 365 para empresas o Exchange Online|Complemento nativo para Android, explorador no aplicable|
|OWA para iPad|iPad 2 o posterior|iOS 6 o posterior|En la última actualización de Office 365 para empresas o Exchange Online|Complemento nativo para iOS, explorador no aplicable|
|OWA para iPhone|iPhone 4S o posterior|iOS 6 o posterior|En la última actualización de Office 365 para empresas o Exchange Online|Complemento nativo para iOS, explorador no aplicable|
|Outlook Web App|iPhone 4, iPad 2, iPod Touch 4 o versiones posteriores|iOS 5 o posterior|En Office 365, Exchange Online o localmente en Exchange Server 2013 o posterior|Safari|


## <a name="additional-resources"></a>Recursos adicionales

- [Office Add-ins platform overview (Información general sobre la plataforma de complementos para Office)](../../docs/overview/office-add-ins.md)
- [Disponibilidad de plataformas y hosts de los complementos de Office](http://dev.office.com/add-in-availability)

