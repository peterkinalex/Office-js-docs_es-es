
# <a name="get-started-with-labsjs-for-office-mix"></a>Introducción a LabsJS para Office Mix



El contenido de LabsJS expone una API (labs.js), ejemplos, documentación y archivos asociados que puede usar para desarrollar laboratorios interactivos, integrarlos en Aplicaciones para Office Mix y representarlos en Microsoft PowerPoint. En realidad, estos laboratorios son Complementos de Office que se crean con HTML5 y la biblioteca de JavaScript de labs.js.

## <a name="labsjs-content"></a>Contenido de LabsJS

LabsJS proporciona documentación, laboratorios de ejemplo y los archivos necesarios para crear y publicar sus propios laboratorios de Aplicaciones para Office Mix.


**Archivos necesarios**


|**Archivo**|**Descripción**|
|:-----|:-----|
|labs-1.0.4.js|La API de JavaScript de LabsJS para el desarrollo de Office Mix Labs. Este archivo se tiene que incluir en el proyecto para permitir que se integre con Office Mix. El archivo también está hospedado en una red de entrega de contenido (CDN) en <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>. Al publicar la aplicación, es necesario vincularla al archivo de la red CDN.|
|labs-1.0.4.d.ts|Archivo de definición TypeScript para labs.js. Esto permite integrar fácilmente el código TypeScript con labs.js. El archivo de definición también proporciona un amplio panorama de todos los componentes que se incluyen en labs.js. Puede descargar TypeScript en [http://www.typescriptlang.org/](http://www.typescriptlang.org/). El archivo de definición se creó con la versión 0.9.1.1 de TypeScript.|
|Historial|Historial de versiones para las diversas versiones de la biblioteca de labs.js.|
|Labshost.html|Una página web que permite ver y depurar el laboratorio con Aplicaciones para Office Mix, fuera del contexto de PowerPoint. Para usar esta página, escriba su dirección URL en el cuadro de entrada principal y se cargará dentro del marco. Los datos intercambiados entre la API y Aplicaciones para Office Mix al ejecutarse en PowerPoint o el reproductor de lecciones de Aplicaciones para Office Mix aparecerán en los cuadros de entrada a la derecha. Los datos también se pueden preinicializar. Tenga en cuenta que los laboratorios de ejemplo en la sección de ejemplos muestran Office Mix Add-ins existentes que se ejecutan en el contexto de host.|
|SampleManifest.xml|Un ejemplo de manifiesto de Complementos de Office para usar como plantilla para crear su propio manifiesto de aplicación.|
|Simplelab.html|Un laboratorio de ejemplo creado con labs.js. Permite la selección y la inserción de una página web, que después realiza un seguimiento del usuario que la está viendo.|
|Simplelab.ts|El archivo TypeScript usado para crear el ejemplo de Simplelab.|
|Simplelab.js|La versión de JavaScript del ejemplo de Simplelab. Este y el archivo simplelab.ts muestran el uso de la API LabsJS.|

## <a name="set-up-your-development-environment"></a>Configurar el entorno de desarrollo

La biblioteca de labs.js sirve como capa de abstracción en la biblioteca office.js (la API de Complementos de Office), por lo que los laboratorios creados con la biblioteca de labs.js son en realidad Complementos de Office. Para poder trabajar con la biblioteca de labs.js y ejecutar estos laboratorios dentro de Aplicaciones para Office Mix, primero debe configurar su cuenta como de desarrollador de Complementos de Office.


### <a name="register-for-an-office-365-developer-site"></a>Registrarse en un sitio para desarrolladores de Office 365

El primer paso es registrarse en un Sitio para desarrolladores de Office 365. Esto permite hospedar y probar el laboratorio antes de enviarlo a la Tienda Office. El sitio permite publicar el complemento en Aplicaciones para Office Mix y probarlo en un entorno activo.

Para obtener más información, vea [Configurar un entorno de desarrollo para complementos para SharePoint en Office 365](http://msdn.microsoft.com/library/b22ce52a-ae9e-4831-9b68-c9210af6dc54%28Office.15%29.aspx). 


### <a name="set-up-an-app-catalog-on-sharepoint-online"></a>Configurar un catálogo de aplicaciones en SharePoint Online

Una vez que se crea y aprovisiona el sitio para desarrolladores, debe configurar un catálogo de complemento en SharePoint Online. Para obtener más información, consulte [Configurar un catálogo de complementos en Office 365](../../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

Para Aplicaciones para Office Mix, use un catálogo de complemento para que pueda insertar add-ins de preproducción en una lección y llevar a cabo comprobaciones integrales antes de enviar los laboratorios a la tienda.


## <a name="create-your-lab"></a>Crear el laboratorio

Para crear su primer laboratorio, siga los pasos del [tutorial](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md), que explica cómo crear un simple cuestionario de verdadero o falso. Consulte [Tutorial: Crear su primer laboratorio para Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md).


## <a name="publish-your-lab"></a>Publicar el laboratorio

Después de crear el laboratorio, puede publicarlo y enviarlo a la tienda.


### <a name="create-and-upload-your-application-manifest"></a>Crear y cargar el manifiesto de aplicación

El manifiesto de aplicación es un documento XML que describe el laboratorio LabJS. Proporciona una referencia a la dirección URL donde se hospeda el laboratorio y detalles sobre este, como el nombre para mostrar, la descripción, los iconos, el tamaño, etc.

Incluimos un manifiesto de ejemplo, "SampleManifest.xml". Para obtener más información sobre el esquema de manifiesto, así como un vínculo a la definición de esquema, consulte [Manifiesto XML de complementos para Office](../../../docs/overview/add-in-manifests.md).

Para cargar el manifiesto en su sitio de SharePoint, vaya primero al catálogo de aplicaciones, que normalmente se encuentra en la dirección URL <code>https://\<your site\>/sites/AppCatalog</code>. Después, elija el botón **Nueva aplicación** y siga los pasos para cargar el manifiesto de aplicación.


### <a name="update-your-powerpoint-2013-catalog"></a>Actualizar el catálogo de PowerPoint 2013

Después actualice el catálogo de PowerPoint 2013. Luego puede iniciar sesión con su cuenta de desarrollador.

Empiece por actualizar el catálogo de PowerPoint 2013. Inicie PowerPoint 2013 y navegue por la ruta del menú  **Archivo > Opciones > Centro de confianza > Configuración del Centro de confianza > Catálogos de aplicaciones de confianza**. Desde ahí, agregue una referencia al catálogo de aplicaciones y elija  **Aceptar**. PowerPoint 2013 le pedirá que cierre sesión para que los cambios surtan efecto. Cierre sesión.

Por último, vuelva a iniciar sesión con la cuenta de desarrollador. Elija el nombre de inicio de sesión en la esquina superior derecha en PowerPoint 2013 e inicie sesión con su cuenta de desarrollador. Ahora puede insertar el complemento.


### <a name="insert-publish-and-view-your-app"></a>Insertar, publicar y ver la aplicación

Para insertar el complemento en el catálogo, elija la cinta  **Insertar** y luego elija **Tienda** en la sección **Aplicaciones**. Elija  **Mi organización** y verá el complemento en el catálogo de complemento. Elija el complemento, seleccione **Insertar** y complemento (laboratorio) se inserta en el documento de PowerPoint 2013.

Ahora puede sacar partido de todas las funciones disponibles en Aplicaciones para Office Mix para publicar la lección con su nuevo laboratorio.


 >**Importante:** Para ver la aplicación, tiene que iniciar sesión en el catálogo de SharePoint con el mismo explorador donde ve la lección. Los catálogos de SharePoint solo permiten el acceso a los usuarios autenticados y, por lo tanto, para poder ver la aplicación primero tiene que iniciar sesión. 


### <a name="submit-your-lab-to-the-office-store"></a>Enviar el laboratorio a la Tienda Office

Para enviar el laboratorio a la Tienda Office, vea [Publicar el complemento de Office](../../publish/publish.md)


## <a name="additional-resources"></a>Recursos adicionales



- [Complementos de Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Complementos de Office](../../../docs/overview/office-add-ins.md)
    
- [Crear el primer laboratorio para Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)
    
