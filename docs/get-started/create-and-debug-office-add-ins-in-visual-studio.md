
# Crear y depurar complementos de Office en Visual Studio




 >**Nota** Estas instrucciones se basan en Visual Studio 2015. Si usa otra versión de Visual Studio, los procedimientos pueden variar ligeramente.



## Crear un proyecto Complemento de Office en Visual Studio


Para empezar, compruebe que tiene instalado [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx). 


1. En la barra de menús de Visual Studio, seleccione **Archivo**  >  **Nuevo**  >  **Proyecto**.
    
2. En la lista de tipos de proyecto en **Visual C#** o **Visual Basic**, expanda **Office/SharePoint**, elija **Complementos web** y, después, seleccione uno de los proyectos de complemento.  
    
3. Asigne un nombre al proyecto y, después, elija **Aceptar** para crear el proyecto.
    
4. Visual Studio crea el proyecto y sus archivos aparecen en el **Explorador de soluciones**. La página predeterminada Home.html se abre en Visual Studio.
    
En Visual Studio 2015, algunas de las plantillas de proyecto del complemento se actualizaron para reflejar funcionalidad adicional:


- Los complementos de contenido pueden aparecer en el cuerpo de documentos Access y PowerPoint y en hojas de cálculo de Excel. También puede seleccionar la opción Proyecto básico para crear un proyecto con complemento de contenido básico con códigos de inicio mínimos, o la opción Proyecto de visualización de documento (para Access y Excel únicamente) para crear un complemento de contenido con más funciones, que incluya código de inicio para visualizar los datos y enlazarlos.
    
- Los complementos de Outlook incluyen opciones no solo para incluir el complemento en mensajes de correo o citas, sino también para especificar si el complemento está disponible cuando se redacta o se lee un mensaje de correo o cita.
    

 >**Nota**  En Visual Studio, la mayoría de las opciones se comprende a partir de sus descripciones, salvo la casilla  **Mensaje de correo electrónico**. Use esa casilla si desea crear un complemento de Outlook que aparezca no solo con elementos de correo, sino también con solicitudes de reuniones, respuestas y cancelaciones.

Una vez que haya completado el asistente, Visual Studio crea una solución que contiene dos proyectos.



|**Project**|**Descripción**|
|:-----|:-----|
|Proyecto de complemento|Contiene solo un archivo de manifiesto XML con todas las configuraciones que describen su complemento. Estas configuraciones ayudan al host de Office a determinar cuándo el complemento se debe activar y dónde debe aparecer. Visual Studio genera los contenidos de esta archivo para que usted pueda ejecutar el proyecto y usar el complemento inmediatamente. Puede cambiar estas configuraciones en cualquier momento usando el editor de manifiesto.|
|Proyecto de aplicación web|Contiene las páginas de contenido de su complemento, incluidos todos los archivos y referencias de archivos que necesita para desarrollar HTML compatible con Office y páginas de JavaScript. Mientras desarrolla su complemento, Visual Studio aloja la aplicación web en su servidor local IIS. Cuando esté listo para publicar, deberá encontrar un servidor para alojar este proyecto.Para obtener más información sobre proyectos de aplicaciones web ASP.NET, vea [Proyectos web de ASP.NET](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).|

## Modificar las configuraciones de su complemento


Para modificar las configuraciones de su complemento, edite el archivo de manifiesto XML del proyecto. En el **Explorador de soluciones**, expanda el nodo del proyecto del complemento, expanda la carpeta que contiene el manifiesto XML y elija el manifiesto XML. Puede apuntar a cualquier elemento en el archivo para ver una información sobre herramientas que describe la finalidad del elemento. Para más información sobre el archivo de manifiesto, vea [Manifiesto XML de complementos de Office](../../docs/overview/add-in-manifests.md).


## Desarrolle el contenido de su complemento


Mientras que el proyecto del complemento le permite modificar las configuraciones que describen su complemento, la aplicación web proporciona el contenido que aparece en el mismo. 

El proyecto de aplicación web contiene una página HTML predeterminada y un archivo de JavaScript que puede usar para empezar. El proyecto también contiene un archivo de JavaScript que es común a todas las páginas que agrega a su proyecto. Estos archivos son convenientes porque contienen referencias a otras bibliotecas de JavaScript, entre ellas la API de JavaScript para Office. 

A medida que su complemento se vuelve más sofisticado, puede agregar más archivos HTML y JavaScript. Puede usar los contenidos de los archivos HTML y JavaScript predeterminados como ejemplos de los tipos de referencias que puede querer agregar a otras páginas en su proyecto para que funcionen con su complemento. La siguiente tabla describe archivos HTML y JavaScript predeterminados.



|**Archivo**|**Descripción**|
|:-----|:-----|
|**Home.html**|Esta es la página HTML predeterminada del complemento y se ubica en la carpeta  **Home** del proyecto. Esta página aparece como la primera dentro del complemento cuando se activa en un documento, mensaje de correo electrónico o elemento de cita. Este archivo es conveniente porque contiene todas las referencias del archivo que necesita para empezar. Cuando esté listo para crear su primer complemento, solo tiene que agregar su código HTML a este archivo.|
|**Home.js**|Este es el archivo JavaScript asociado con la página Home.js y se ubica en la carpeta  **Home** de este proyecto. Puede colocar cualquier código específico del comportamiento de la página Home.html en el archivo Home.js. El archivo Home.js contiene algunos códigos de ejemplo para que empiece.|
|**App.js**|Este es el archivo de JavaScript predeterminado de todo el complemento y se ubica en la carpeta  **Complemento** del proyecto. Puede colocar un código común al comportamiento de múltiples páginas de su complemento en el archivo App.js. El archivo App.js contiene algunos códigos de ejemplo para que empiece.|

 >**Nota**  No tiene que usar estos archivos. Puede agregar otros archivos al proyecto y usarlos en su lugar. Si desea que aparezca otro archivo HTML como página inicial del complemento, abra el editor de manifiesto y luego señale la propiedad  **SourceLocation** en el nombre del archivo.


## Depurar el complemento


Cuando esté listo para iniciar el complemento, revise la compilación y depure las propiedades relacionadas y luego inicie la solución.


### Revisar las propiedades de compilación y depuración

Antes de iniciar la solución, es buena idea asegurarse de que Visual Studio abra la aplicación host que usted quiere. Esa información aparece en las páginas de propiedades del proyecto junto con varias otras propiedades que se refieren a la compilación y la depuración del complemento.


### Para abrir las páginas de propiedades de un proyecto


1. En  **Explorador de soluciones**, elija el nombre del proyecto.
    
2. En la barra de menús, elija  **Ver**,  **Ventana Propiedades**.
    
En la tabla siguiente se describen las propiedades del proyecto.



|**Propiedad**|**Descripción**|
|:-----|:-----|
|**Acción de inicio**|Especifica si se debe depurar el complemento en un cliente de escritorio de Office o en un cliente de Office Online en el explorador especificado.|
|**Documento de inicio** (solo complementos de panel de tareas y de contenido)|Especifica qué documento debe abrirse cuando inicia el proyecto.|
|**Proyecto web**|Especifica el nombre del proyecto web asociado con el complemento.|
|**Dirección de correo electrónico** (solo complementos de Outlook)|Especifica la dirección de correo electrónico de la cuenta de usuario en Exchange Server o Exchange Online con la que quiere probar el complemento de Outlook.|
|**Dirección URL de EWS** (solo complementos de Outlook)|Dirección URL del servicio web de Exchange (por ejemplo: https://www.contoso.com/ews/exchange.aspx). |
|**Dirección URL de OWA** (solo complementos de Outlook)|Dirección URL de Outlook Web App (por ejemplo: https://www.contoso.com/owa).|
|**Nombre de usuario** (solo complemento de Outlook)|Especifica el nombre de la cuenta de usuario en Exchange Server o Exchange Online.|
|**Archivo de proyecto**|Especifica el nombre del archivo que contiene la compilación, configuración y otra información sobre el proyecto.|
|**Carpeta de proyecto**|La ubicación del archivo de proyecto.|

### Usar un documento existente para depurar el complemento (solo complementos de panel de tareas y de contenido)


Puede agregar documentos al proyecto del complemento. Si tiene un documento que contiene datos de prueba que quiere usar con la aplicación, Visual Studio lo abrirá cuando se inicie el proyecto.


### Para usar un documento existente para depurar el complemento


1. En el  **Explorador de soluciones**, elija la carpeta del proyecto del complemento.
    
     >**Nota** Elija el proyecto de complemento y no el proyecto de aplicación web.
2. En el menú  **Proyecto**, elija  **Agregar elemento existente**.
    
3. En el cuadro de diálogo  **Agregar elemento existente**, busque y seleccione el documento que quiere agregar.
    
4. Elija el botón  **Agregar** para agregar el documento al proyecto.
    
5. En el  **Explorador de soluciones**, abra el menú contextual del proyecto y luego elija  **Propiedades**.
    
    Aparecen las páginas de propiedades del proyecto.
    
6. En la lista  **Documento de inicio**, elija el documento que agregó al proyecto y, a continuación, elija el botón  **Aceptar** para cerrar las páginas de propiedades.
    

### Iniciar la solución


Visual Studio compilará automáticamente la solución cuando la inicie. Puede iniciar la solución desde la barra  **Menú** mediante la elección de **Depurar**,  **Iniciar**. 


 >**Nota**  Si la depuración de scripts no está habilitada en Internet Explorer, no podrá iniciar el depurador en Visual Studio. Puede habilitar la depuración de scripts si abre el cuadro de diálogo  **Opciones de Internet**, selecciona la pestaña  **Opciones avanzadas** y desactiva las casillas **Deshabilitar la depuración de scripts (Internet Explorer)** y **Deshabilitar la depuración de scripts (otros)**.

Visual Studio compila el proyecto y lleva a cabo las siguientes acciones.


1. Crea una copia del archivo de manifiesto XML y la agrega al directorio  _ProjectName_\Output. La aplicación host usa esta copia cuando se inicia Visual Studio y se depura el complemento.
    
2. Crea un conjunto de entradas de registro en el equipo que permiten que el complemento aparezca en la aplicación host.
    
3. Compila el proyecto de aplicación web y lo implementa en el servidor web IIS local (http://localhost). 
    
Luego, Visual Studio lleva a cabo las acciones siguientes.


1. Modifica el elemento [SourceLocation](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx) del archivo de manifiesto XML al reemplazar el token ~remoteAppUrl con la dirección completa de la página de inicio (por ejemplo, http://localhost/MyAgave.html).
    
2. Inicia el proyecto de aplicación web en IIS Express.
    
3. Abre la aplicación host. 
    
Visual Studio no muestra errores de validación en la ventana  **OUTPUT** cuando se compila el proyecto. Visual Studio informa de los errores y avisos en la ventana **ERRORLIST** cuando se generan. Visual Studio también informa de los errores de validación con un subrayado ondulado de diferentes colores en el editor de código y de textos. Estas marcas le informan de los problemas que Visual Studio ha detectado en su código. Para más información, vea [Código y editor de texto](http://go.microsoft.com/fwlink/?LinkID=128497). Para más información sobre cómo habilitar o deshabilitar la validación, vea los temas siguientes: 


- [Opciones, editor de texto, JavaScript, IntelliSense](http://go.microsoft.com/fwlink/?LinkID=238779)
    
- [Cómo: Establecer opciones de validación para la edición de HTML en Visual Web Developer](http://msdn.microsoft.com/en-us/library/vstudio/0byxkfet%28v=vs.100%29.aspx)
    
- [Validación, CSS, Editor de texto, Opciones (Cuadro de diálogo)](http://go.microsoft.com/fwlink/?LinkID=238780)
    
Para revisar las reglas de validación del archivo de manifiesto XML en su proyecto, vea [Manifiesto XML de complementos para Office](../../docs/overview/add-in-manifests.md).


### Mostrar una aplicación en Excel, Word o Project y revisar el código


Si establece la propiedad  **Documento de inicio** del proyecto de complemento en Excel o Word, Visual Studio crea un documento nuevo y aparece el complemento. Si establece la propiedad **Documento de inicio** del proyecto de complemento para que use un documento existente, Visual Studio abre el documento, pero debe insertar el complemento manualmente. Si establece el **Documento de inicio** como **Microsoft Project**, también deberá insertar el complemento de forma manual.


### Para mostrar una Complemento de Office en Excel o Word


1. En Excel o Word, en la pestaña  **Insertar**, elija  **Aplicaciones para Office**.
    
2. En la lista que aparece, elija su complemento.
    

### Para mostrar una Complemento de Office en Project


1. En Project, en la pestaña  **Project**, elija  **Complementos de Office**.
    
2. En la lista que aparece, elija su complemento.
    
A continuación, en Visual Studio, puede configurar puntos de interrupción. Luego, a medida que interactúe con su complemento, podrá repasar el código en sus archivos de código HTML, JavaScript y C# o VB.


### Mostrar el complemento de Outlook en Outlook y repasar el código


Para ver la aplicación en Outlook, abra un mensaje de correo electrónico o un elemento de cita.

Outlook activa la complemento para el elemento siempre que se cumplan los criterios de activación. La barra de la complemento aparece en la parte superior de la ventana del inspector o el panel de lectura, y el complemento de Outlook aparece como un botón en la barra de la complemento. Si el complemento tiene un comando de complemento, aparecerá un botón en la cinta de opciones, ya sea en la pestaña predeterminada o en una pestaña personalizada especificada, y el complemento no aparecerá en la barra del complemento.

Para ver el complemento de Outlook, elija el botón para su complemento de Outlook.

En Visual Studio, puede configurar puntos de interrupción. Luego, mientras interactúe con su complemento de Outlook, podrá repasar el código en su HTML, JavaScript o en los archivos de código C# o VB. 

También puede cambiar el código y ver los efectos de esos cambios en el complemento de Outlook sin tener que cerrar el Complemento de Office e iniciar el proyecto de nuevo. En Outlook, simplemente abra el menú contextual para el complemento de Outlook y luego seleccione  **Volver a cargar**.


### Modificar código y seguir depurando el complemento sin tener que volver a iniciar el proyecto


Puede cambiar el código y ver los efectos de esos cambios en su aplicación sin tener que cerrar la aplicación host e iniciar el proyecto de nuevo. Después de cambiar el código, abra el menú contextual de la aplicación y, a continuación, elija  **Recargar**. Al volver a cargar la aplicación, se desconecta del depurador de Visual Studio. Por ello puede ver los efectos del cambio, pero no podrá volver a revisar el código paso por paso hasta que asocie el depurador de Visual Studio a todos los procesos de iexplore.exe disponibles.


### Para asociar el depurador de Visual Studio a todos los procesos de Iexplore.exe disponibles


1. En Visual Studio, elija  **DEPURAR**,  **Asociar al proceso**.
    
2. En el cuadro de diálogo  **Asociar al proceso**, elija todos los procesos disponibles de  **Iexplore.exe** y, a continuación, elija el botón **Asociar**.
    

## Pasos siguientes

- [Publicar el complemento para Office](../publish/publish.md)
    
