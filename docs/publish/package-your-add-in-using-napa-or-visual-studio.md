
# <a name="package-your-add-in-using-napa-or-visual-studio-to-prepare-for-publishing"></a>Empaquetar el complemento con Napa o Visual Studio para prepararlo para su publicación

El complemento de Office contiene un archivo XML que se usará para publicar el complemento. Deberá publicar los archivos de la aplicación web del proyecto de forma independiente.

## <a name="package-an-office-add-in-that-you-create-by-using-napa"></a>Empaquetar una Complemento de Office creada con Napa



1. En Napa, en el lateral de la página, elija el botón  **Publicar** ( ![Botón Publicar](../../images/Apps_NAPA_Publish.png)).
    
2. En el cuadro de diálogo  **Configuración de publicación**, elija el botón  **Siguiente**.
    
3. Proporcione la URL del sitio web que hospedará los archivos de contenido de su complemento (por ejemplo, los archivos HTML y JavaScript predeterminados del proyecto) y luego elija el botón  **Publicar**.
    
4. En el cuadro de diálogo  **Publicación correcta**, elija el vínculo  **Ubicación de publicación**.
    
    Aparecerá una biblioteca de documentos que contiene el archivo de manifiesto XML de complemento y los archivos de contenido web. 
    
Luego, manualmente, copie los archivos de contenido web (hojas de estilo, archivos de JavaScript y archivos HTML) en el servidor web que hospeda el sitio web que proporcionó en el cuadro de diálogo  **Configuración de publicación**.

Ahora puede cargar el manifiesto XML en la ubicación adecuada para [publicar el complemento](../publish/publish.md). 


## <a name="deploy-your-web-project-and-package-your-add-in-by-using-visual-studio-2015"></a>Implementar el proyecto web y empaquetar el complemento mediante Visual Studio 2015



### <a name="to-deploy-your-web-project"></a>Para implementar su proyecto web


1. En el  **Explorador de soluciones**, abra el menú contextual para el proyecto de complemento y luego elija  **Publicar**.
    
    Aparecerá la página **Publicación de los complementos**.
    
2. En la lista desplegable **Perfil actual**, seleccione un perfil o elija **Nuevo…** para crear un perfil.
    
     >**Nota** Un perfil de publicación especifica el servidor en que se va a realizar la implementación, las credenciales necesarias para iniciar sesión en el servidor, las bases de datos que se implementarán y otras opciones de implementación.

    Si elige **Nuevo...**, aparece el asistente para **Crear perfil de publicación**. Puede usar a este asistente para importar un perfil de publicación desde un proveedor de hospedaje de sitios web como Microsoft Azure o crear un nuevo perfil y agregar el servidor, las credenciales y otras opciones en el procedimiento siguiente.
    
    Para obtener más información acerca de la importación de perfiles de publicación o crear nuevos perfiles de publicación, consulte [Creación de un perfil de publicación](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).
    
3. En la página **Publique el complemento**, elija el vínculo **Implemente su proyecto web**.
    
    Se abrirá el cuadro de diálogo **Publicación web**. Para obtener más información acerca de cómo usar este asistente, consulte [Procedimiento para implementar un proyecto web mediante la publicación con clic en Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).
    

### <a name="to-package-your-add-in"></a>Para empaquetar el complemento


1. En la página  **Publique el complemento**, elija el vínculo  **Empaquetar la aplicación**.
    
    Aparecerá el asistente **Publicar complementos para Office y SharePoint**.
    
2. En la lista desplegable  **¿Dónde está hospedado el sitio web?**, seleccione o escriba una URL del sitio web que hospedará los archivos de contenido de su complemento y luego elija el botón  **Finalizar**.
    
    Tendrá que especificar una dirección que comience con el prefijo HTTPS para completar este asistente. En general, usar un extremo HTTPS para el sitio web es el mejor método, pero no es necesario si no planea publicar el complemento en la Tienda Office. Una vez creado el paquete, puede abrir el manifiesto en el Bloc de notas y reemplazar el prefijo HTTPS del sitio web con un prefijo HTTP. Para obtener más información, consulte [¿Por qué los complementos deben tener seguridad SSL?](http://msdn.microsoft.com/en-us/library/jj591603#bk_q7). 
    
     >**Nota** Los sitios web de Azure proporcionan automáticamente un extremo HTTPS.

    Visual Studio genera los archivos que necesita para publicar el complemento y, después, abre la carpeta de salida de publicación. 
    
Si planea enviar el complemento a la Tienda Office, haga clic en el vínculo  **Realice una comprobación de validación** para identificar los problemas que impedirán la aceptación del complemento. Debe resolver todos los problemas antes de enviar el complemento a la tienda.

Ahora puede cargar el manifiesto XML en la ubicación adecuada para [publicar el complemento](../publish/publish.md). Puede encontrar el manifiesto XML en  `OfficeAppManifests` en la carpeta `app.publish`. Por ejemplo:

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="additional-resources"></a>Recursos adicionales



- [Publicar el complemento de Office](../publish/publish.md)
    
- 
  [Enviar complementos de Office y SharePoint, y aplicaciones web de Office 365 a la Tienda Office](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
