# <a name="design-guidelines-for-office-add-ins"></a>Directrices de diseño para complementos de Office

Los Complementos de Office amplían la experiencia de Office proporcionando funcionalidad contextual a la que pueden acceder los usuarios desde clientes de Office. Los complementos permiten a los usuarios hacer más cosas, ya que habilitan el acceso a funcionalidades de terceros desde Office, sin cambios de contexto costosos. 

 El diseño de la experiencia de usuario del complemento debe integrarse a la perfección con Office para proporcionar una interacción eficaz y natural para los usuarios. Aproveche los comandos de complementos (extensiones de la interfaz de usuario de Office) para proporcionar acceso a su complemento, y use los [elementos de la interfaz de usuario](ui-elements/ui-elements.md) y las [prácticas recomendadas](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices) que le sugerimos al crear la interfaz de usuario personalizada basada en HTML. 
 
 
## <a name="core-office-add-in-design-principles"></a>Principios de diseño esenciales para los Complementos de Office 
Independientemente del marco de trabajo subyacente que use para crear la interfaz de usuario personalizada, aplique los siguientes principios al diseñar el complemento: 

- **Diseñe explícitamente para Office**. La funcionalidad y el aspecto de un complemento deben concordar armoniosamente con la experiencia de Office, incluida la aplicación del tema de Office o del documento.
 
- **Haga más eficientes a los usuarios**. Ayude a los usuarios a realizar un trabajo sin que interfiera con el resto de las tareas. Permita una interacción óptima entre los documentos de Office y el complemento. 

- **Dé prioridad al contenido sobre el aspecto**. Haga hincapié en el contenido y la funcionalidad del complemento sobre el aspecto. Aproveche al máximo el espacio, evitando los elementos superfluos de la interfaz de usuario que no agregan valor a la experiencia del usuario.  

- **Mantenga a los usuarios al mando**. Permita a los usuarios controlar su experiencia, entender las decisiones importantes y revertir fácilmente las acciones que realiza el complemento. 

- 
  **Diseñe para todas las plataformas y métodos de entrada**. Los complementos están diseñados para que funcionen en todas las plataformas que admite Office, y la experiencia de usuario de su complemento debe estar optimizada para que funcione en diversas plataformas y factores de forma. Permita la compatibilidad con dispositivos dotados de ratón/teclado y entrada táctil, y asegúrese de que la interfaz de usuario HTML personalizada tiene capacidad de respuesta para adaptarse a diferentes factores de forma. Para obtener más información, consulte [Tecnología táctil](https://msdn.microsoft.com/EN-US/library/mt590883.aspx#bk_Touch). 


## <a name="design-language"></a>Lenguaje de diseño
Le recomendamos que adopte el lenguaje de diseño de Office y que use [Office UI Fabric](https://dev.office.com/fabric) para crear experiencias personalizadas basadas en HTML en los complementos. Si su organización ya cuenta con un lenguaje de diseño, puede usarlo, pero siempre que el resultado final sea una experiencia armoniosa para los usuarios de Office. 


## <a name="add-in-building-blocks"></a>Bloques de creación de complementos
Puede usar dos tipos de elementos de la interfaz de usuario para crear sus complementos: 

- Los [comandos de complemento](ui-elements/ui-elements.md#add-in-commands) le permiten agregar enlaces nativos de experiencia del usuario a las aplicaciones de Office
- La [interfaz de usuario personalizada basada en HTML](ui-elements/ui-elements.md#custom-html-based-ui) permite aprovechar la potencia de HTML en los clientes de Office. 

Para más información sobre cómo usar estos bloques de creación, vea [Elementos de la interfaz de usuario](ui-elements/ui-elements.md).  

## <a name="ux-design-patterns"></a>Modelos de diseño de la experiencia del usuario

Para ayudarle a crear una experiencia del usuario de primera clase para su complemento, proporcionamos plantillas donde se muestran modelos de diseño de la experiencia del usuario comunes. Estas plantillas reflejan los [procedimientos recomendados](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices) para crear complementos atractivos de primer nivel y, además, se incluyen modelos para experiencias de primera ejecución, elementos de personalización de marca y notificaciones de usuario. Usan componentes y estilos de [Office UI Fabric](https://dev.office.com/fabric), y se incluyen elementos que amplían de forma natural la interfaz de usuario de Office.

Para tener acceso a las plantillas, vea el repositorio [Modelos de diseño de la experiencia del usuario para complementos de Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns). También hay disponibles archivos de Adobe Illustrator, que puede descargar y actualizar para reflejar sus propios diseños. También puede copiar los archivos de código del repositorio [Modelos de diseño de la experiencia del usuario para complementos de Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) en su proyecto de complemento y personalizarlos según sea necesario. 

## <a name="recommended-layouts-and-interaction-patterns"></a>Diseños y patrones de interacción recomendados
Ofrecemos diseños recomendados para cada tipo de complemento, junto con ejemplos **completos** que le ayudarán a combinar todos los elementos. Para obtener más información sobre cómo diseñar un complemento, consulte lo siguiente:

- [Diseño para contenedores de panel de tareas](ui-elements/layout-for-task-pane-add-ins.md)
- [Diseños para complementos de contenido](ui-elements/layout-for-content-add-ins.md) 
- [Diseños para complementos de correo](ui-elements/layouts-for-outlook-add-ins.md)

Consulte también Patrones de interacción para obtener ejemplos de escenarios comunes de complementos y sus patrones de interacción correspondientes.

## <a name="additional-resources"></a>Recursos adicionales

- [Office UI Fabric](https://dev.office.com/fabric) 

