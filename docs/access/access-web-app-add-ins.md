
# Crear complementos para aplicaciones web de Access



Este artículo explica cómo usar Visual Studio 2015 para desarrollar un complemento de Office dirigido a aplicaciones web de Access.

>
  **Nota:** Para obtener información sobre cómo desarrollar soluciones para Access con VBA, consulte [Access](https://msdn.microsoft.com/en-us/library/fp179695.aspx) en MSDN.

## Requisitos previos

Para crear una Complemento de Office dirigida a aplicaciones web de Access, necesita:


- Visual Studio 2015

- Un sitio de SharePoint Online (incluido en muchas suscripciones a Office 365). Este sitio debe tener un catálogo de complementos. Para más información, vea [Procedimiento para configurar un catálogo de complementos en SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).


 >**Nota**  Complementos de Office trabajará con aplicaciones web de Access alojado en SharePoint Online o Office 365. La aplicación de escritorio Access 2013 no es compatible con Complementos de Office. Complementos de Office dirigidas a aplicaciones web de Access son compatibles con la version 1.1 y posteriores de Office.js.


## Crear un proyecto de complemento para Access en Visual Studio


1.  Abra Visual Studio y, en el menú, elija **Archivo**, **Nuevo**, **Proyecto**. Se abrirá el cuadro de diálogo **Nuevo proyecto**.

2. En el cuadro de diálogo  **Nuevo proyecto**, en el panel izquierdo, vaya a  **Instalado**,  **Plantillas**,  **Visual C#**,  **Office/SharePoint**,  **Complementos de Office**.

3. En el cuadro de diálogo  **Nuevo proyecto**, elija en el panel central  **Complemento de Office**.

4. En la parte inferior del cuadro de diálogo, escriba un nombre para su proyecto y elija  **Aceptar**. Se abrirá el cuadro de diálogo  **Crear complemento para Office**.

5. En el cuadro de diálogo  **Crear complemento de Office**, elija  **Contenido** y luego **Siguiente**.

6. En la pantalla siguiente del cuadro de diálogo  **Crear complemento para Office**, elija  **Complemento básico** o **Aplicación de visualización de documentos** y asegúrese de que la casilla **Access** esté activada.

7. Una vez listo, elija  **Finalizar**. Visual Studio creará un proyecto de arranque sobre el que puede basarse.

8. En el  **Explorador de soluciones**, elija el proyecto web del proyecto ( **project_name>Web**). En el panel de propiedades, busque la entrada para  **SSL URL**. Debería ser similar a:  `https://localhost:44314/`. Seleccione esta dirección URL y cópiela a su portapapeles. La necesitará pronto.

9. Haga clic con el botón derecho en el nombre del proyecto en el **Explorador de soluciones**. En el menú contextual, elija **Publicar**. Se abrirá el asistente **Publique el complemento**.

10. En el asistente **Publique el complemento**, seleccione la lista desplegable junto a **Perfil actual**. En esta lista desplegable, elija **Nuevo**. Se abrirá el cuadro de diálogo **Publicar complementos de Office y SharePoint**.

11. En este cuadro de diálogo, elija  **Crear perfil nuevo**, escriba un nombre reconocible para el perfil y luego elija  **Finalizar**. Se cerrará el cuadro de diálogo  **Publicar complementos para Office y SharePoint** y se regresará al asistente **Publique el complemento**.

12. En el asistente, elija  **Empaquetar el complemento**. Esto finalizará el complemento para que pueda publicarse en un catálogo de complementos en SharePoint.

13. En la siguiente página, para  **¿Dónde está hospedado el sitio web?** coloque la dirección URL para el host del sitio web. Puede ser el valor de **Dirección URL de SSL** que copió en el paso 8. Después, elija **Finalizar**.

14. En el  **Explorador de soluciones**, haga clic con el botón derecho en el nodo de manifiesto del proyecto (directamente debajo del nombre del proyecto) y seleccione  **Abrir carpeta en el Explorador de archivos**. Tome nota de la ruta para llegar a este archivo. Necesitará este valor más adelante.


 >**Nota**  No puede depurar el complemento sin implementarlo con unaAccess web app.


## Revisar el manifiesto y el archivo Home.Html


1. En el proyecto de Visual Studio, abra el archivo  **Home.html** y busque las líneas que hacen referencia a la biblioteca de scripts de office.js.

```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```
 >**Nota:** La etiqueta de script hace referencia a la versión 1.1 (y posteriores) de Office.js. Access usa los elementos de la API introducidos en la versión 1.1.

2. Abra el archivo del manifiesto asociado con su proyecto. Se le asignará el nombre a este archivo después del nombre de su proyecto y tendrá la extensión ".xml".

3.  En el archivo del manifiesto, busque la sección **Hosts** y busque una entrada de **Host**.

```xml
  <Hosts> <Host Name="Database" /> </Hosts>
```
 >**Nota:** Aquí es donde se enumeran las aplicaciones que pueden usar el complemento. Como seleccionó **Access** en el cuadro de diálogo **Crear complemento de Office**, se muestra **Base de datos**. Si incluyó Excel, también verá una entrada para **Libro**.

Las Complementos de Office y SharePoint se basan en la Web. Esto significa que el código del complemento debe hospedarse en un servidor web. Para este ejemplo, el servidor web es su equipo de desarrollo. El servidor debe estar en ejecución para poder probar el complemento. En este caso, eso significa que Visual Studio debe ejecutar el complemento al mismo tiempo que ve y depura el complemento dentro de SharePoint.

Para que un usuario de encuentre y use un complemente, debe estar registrado en el Catálogo de complementos en SharePoint. La información que necesita el Catálogo de complementos está contenida en el archivo de manifiesto.

 >**Nota**  Deberá crear una Access web app para hospedar su Complemento de Office.


## Publicar el complemento en un catálogo de SharePoint Online


1.  Inicie sesión en SharePoint Online o Office 365 y luego vaya al **centro de administración de SharePoint** eligiendo **Administración** en la barra de herramientas de Office 365 en la perte superior de la página.

2. En la página  **Centro de administración de SharePoint**, en la barra de vínculos que está a la izquierda, elija  **Complementos**. Esto lo llevará a la vista de complementos.

3. En el panel central de la página, elija  **Catálogo de complementos**. Se abrirá la página  **Catálogo**.

4. En la página  **Catálogo**, elija  **Distribución de aplicaciones de Office**. Esto lo llevará a una página de directorio llamada  **Complementos para Office** que contiene todas las Complementos de Office instaladas.

5. En la parte superior de la página  **Complementos para Office** elija **Nuevo complemento**. Aparecerá el cuadro de diálogo **Agregar un documento**.

6. En el cuadro de diálogo  **Agregar un documento**, elija  **Examinar** y luego vaya hasta la ubicación del archivo del manifiesto en el proyecto de Visual Studio. Si copió la dirección del archivo de manifiesto antes, puede pegarla en este cuadro de diálogo.

7. Elija el archivo de manifiesto de su proyecto y elija  **Aceptar**. SharePoint agregará ahora el complemento a la biblioteca de SharePoint local.


 >**Nota**  En este procedimiento se supone que ha creado un sitio de prueba para su SharePoint. Si no lo hizo, puede hacerlo desde la pestaña  **Sitios**, en la parte superior de la venta de SharePoint. Puede usar una aplicaciones web de Access existente, si la tiene.


## Cree una Access web app para que aloje a su complemento


1. Vaya hasta el sitio de prueba. En la barra de vínculos que se encuentra a la izquierda, elija  **Contenidos del sitio**. Esto lo llevará a la página  **Contenidos del sitio** de su sitio de prueba.

2. En la página  **Contenidos del sitio**, elija  **Agregar un complemento**. Se abrirá la página  **Contenidos del sitio: Sus aplicaciones**.

3. En la página  **Contenidos del sitio: Sus aplicaciones**, use la barra de búsqueda en la parte superior de la pagina para buscar  **Access Aplicación para Project**.

4. Ahora debe ver un icono para la **Aplicación de Access**.

     >**Nota** Recuerde que no es el complemento de Office, es un nueva aplicación web de Access. Esta aplicación web de Access hospedará el complemento de Office.
5. Al elegir este icono llega al cuadro de diálogo  **Agregar una aplicación de Access**. Escriba un nombre único para su Accessaplicación y seleccione  **Crear**. SharePoint puede tardar un tiempo en crear su aplicación. Al finalizar, verá su Accessaplicación en la página  **Contenidos del sitio** con una etiqueta **Nueva** al lado.

6. Debe abrir la aplicación Accessaplicación en la versión de escritorio de Microsoft Access 2013 y agregar datos antes de que se pueda abrir y ver en SharePoint.


## Agregar su complemento a una aplicaciones web de Access


1. Abra una aplicaciones web de Access.

2. En el complemento, busque en la esquina superior izquierda un icono de engranaje en la barra de pestañas de SharePoint. Elija el engranaje y aparecerá un menú. Elija el elemento de menú  **Complementos para Office**. Se abrirá el cuadro de diálogo  **Comeplementos para Office**.

3. Elija la vista  **Mi organización** y espere a que SharePoint rellene el cuadro de diálogo con las Complementos de Office disponibles.

    Uno de los complementos en el cuadro de diálogo debe ser el complemento de Office que se ha registrado en un procedimiento anterior. Elija ese complemento para insertarlo en la aplicación web de Access. Recuerde que debe ejecutar la aplicación en Visual Studio para poder detectarla y que aparezca en la página de acceso a la aplicación web de Access.


## Depurar el complemento para Office

Para depurar sus complementos, presione F12 en Internet Explorer o elija el icono de engranaje en la barra de tareas del explorador (no el icono de engranaje de la página SharePoint). De esta forma se abren las herramientas de depuración que proporciona Internet Explorer 11. Si usa otro explorador, consulte su documentación para determinar cómo especificar el modo de depuración.

En este punto se pueden establecer los puntos de interrupción, avanzar en los códigos de JavaScript, explorar el DOM y modificar el código para confirmar que los cambios aparecen en la Complemento de Office que está dirigidaaplicaciones web de Access. Consulte [Uso de herramientas para desarrolladores F12](http://msdn.microsoft.com/library/ie/bg182326%28v=vs.85%29) para obtener más información.


## Siguientes pasos

Descargue la muestra [Office 365: enlazar y manipular datos en una aplicación web de Access](https://code.msdn.microsoft.com/officeapps/Office-365-Bind-and-4876274e) para obtener más información sobre cómo implementar una Office Add-in que manipule datos en una Access web app.


## Recursos adicionales



- [Información sobre la API de JavaScript para complementos](../develop/understanding-the-javascript-api-for-office.md)

- [API de JavaScript para Office](../../reference/javascript-api-for-office.md)

