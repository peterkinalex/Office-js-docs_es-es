
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>Empaquetar el complemento con Visual Studio para prepararlo para su publicación

El complemento de Office contiene un archivo XML que se usará para publicar el complemento. Deberá publicar los archivos de la aplicación web del proyecto de forma independiente.


## <a name="deploy-your-web-project-and-package-your-add-in-by-using-visual-studio-2015"></a>Implementar el proyecto web y empaquetar el complemento mediante Visual Studio 2015



### <a name="to-deploy-your-web-project"></a>Para implementar su proyecto web


1. En el  **Explorador de soluciones**, abra el menú contextual para el proyecto de complemento y luego elija  **Publicar**.
    
    Aparecerá la página **Publicación de los complementos**.
    
2. En la lista desplegable **Perfil actual**, seleccione un perfil o elija **Nuevo…** para crear un perfil.
    
     >**Nota** Un perfil de publicación especifica el servidor en que se va a realizar la implementación, las credenciales necesarias para iniciar sesión en el servidor, las bases de datos que se implementarán y otras opciones de implementación.

    Si elige **Nuevo...**, aparece el asistente para **Crear perfil de publicación**. Puede usar a este asistente para importar un perfil de publicación desde un proveedor de hospedaje de sitios web como Microsoft Azure o crear un nuevo perfil y agregar el servidor, las credenciales y otras opciones en el procedimiento siguiente.
    
    Para obtener más información acerca de la importación de perfiles de publicación o crear nuevos perfiles de publicación, consulte [Creación de un perfil de publicación](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).
    
3. En la página  **Publique el complemento**, elija el vínculo  **Implemente su proyecto web**.
    
    The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).
    

### <a name="to-package-your-add-in"></a>Para empaquetar el complemento


1. En la página  **Publique el complemento**, elija el vínculo  **Empaquetar la aplicación**.
    
    Aparecerá el asistente **Publicar complementos para Office y SharePoint**.
    
2. En la lista desplegable  **¿Dónde está hospedado el sitio web?**, seleccione o escriba una URL del sitio web que hospedará los archivos de contenido de su complemento y luego elija el botón  **Finalizar**.
    
    You have to specify an address that begins with the HTTPS prefix to complete this wizard. In general, using an HTTPS endpoint for your website is the best approach, but it is not required if you don't plan to publish your add-in to the Office Store. After the package is created, you can open the manifest in Notepad and replace the HTTPS prefix of your website with an HTTP prefix. For more information, see [Why do my add-ins have to be SSL-secured?](http://msdn.microsoft.com/en-us/library/jj591603#bk_q7). 
    
     >**Nota** Los sitios web de Azure proporcionan automáticamente un extremo HTTPS.

    Visual Studio genera los archivos que necesita para publicar el complemento y, después, abre la carpeta de salida de publicación. 
    
Si planea enviar el complemento a la Tienda Office, haga clic en el vínculo  **Realice una comprobación de validación** para identificar los problemas que impedirán la aceptación del complemento. Debe resolver todos los problemas antes de enviar el complemento a la tienda.

Ahora puede cargar el manifiesto XML en la ubicación adecuada para [publicar el complemento](../publish/publish.md). Puede encontrar el manifiesto XML en  `OfficeAppManifests` en la carpeta `app.publish`. Por ejemplo:

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="additional-resources"></a>Recursos adicionales



- [Publicar el complemento de Office](../publish/publish.md)
    
- [Enviar complementos de Office y SharePoint, y aplicaciones web de Office 365 a la Tienda Office](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
