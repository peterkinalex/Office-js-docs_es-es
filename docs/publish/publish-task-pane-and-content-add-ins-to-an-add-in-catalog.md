
# <a name="publish-task-pane-and-content-add-ins-to-an-add-in-catalog-on-sharepoint"></a>Publicar complementos de panel de tareas y de contenido en un catálogo de complementos de SharePoint

Un catálogo de complementos es una colección de sitios dedicada en una aplicación web de SharePoint o un arrendamiento de SharePoint Online que contiene bibliotecas de documentos para los complementos de Office y SharePoint. Los administradores pueden cargar archivos de manifiesto de complementos de Office en el catálogo de complementos para su uso dentro de su organización. Cuando un administrador registra un catálogo de complementos como catálogo de confianza, los usuarios pueden insertar el complemento desde la interfaz de usuario de inserción en una aplicación cliente de Office.

>**Nota:** Los catálogos de complementos de SharePoint no admiten funciones de complemento que se hayan implementado en el nodo VersionOverrides del [manifiesto del complemento](../overview/add-in-manifests.md).

No se admiten los catálogos de SharePoint en Office 2016 para Mac. Para implementar complementos de Office en clientes Mac, debe enviarlos a la [Tienda Office](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx).   

## <a name="to-set-up-an-add-in-catalog-on-sharepoint"></a>Para configurar un catálogo de complementos en SharePoint

1. Vaya al **Sitio de administración central** (**Inicio** > **Todos los programas** > **Productos de Microsoft SharePoint 2013** > **Administración central de SharePoint 2013**).
    
2. En el panel de tareas de la izquierda, elija  **Complementos**.
    
3. En la página  **Complementos**, vaya a  **Administración de complementos** y elija **Administrar catálogo de complementos**.
    
4. En la página  **Administrar catálogo de complementos**, asegúrese de que se encuentra seleccionada la aplicación web correcta en el  **Selector de aplicaciones web**.
    
5. Elija  **Ver la configuración de sitio**.
    
6. En la página  **Configuración del sitio**, elija  **Administradores de la colección de sitios** para especificar quiénes son estos administradores. A continuación, elija **Aceptar**.
    
7. Para conceder permisos del sitio a los usuarios, elija  **Permisos del sitio** y **Conceder permisos**.
    
8. En el cuadro de diálogo  **Compartir "Sitio del catálogo de aplicaciones"**, especifique uno o varios usuarios del sitio, defina los permisos apropiados para ellos, establezca otras opciones si así lo desea y elija  **Compartir**.
    
9. Para agregar complementos al catálogo de complementos de Complementos de Office, seleccione **Complementos de Office**.

## <a name="to-set-up-an-add-in-catalog-on-office-365"></a>Para configurar un catálogo de complementos en Office 365

1. En la página del centro de administración de Office 365, elija **Administrador** y luego **SharePoint**.
    
2. En el panel de tareas de la izquierda, elija  **complementos**.
    
3. En la página  **complementos**, elija  **Catálogo de complementos**.
    
4. En la página  **Sitio del catálogo de complementos**, elija  **Aceptar** para aceptar la opción predeterminada y crear un nuevo sitio del catálogo de complementos.
    
5. En la página  **Crear colección de sitios del catálogo de complementos**, especifique el título del sitio del catálogo de complementos.
    
6. Especifique la dirección del sitio web.
    
7. Establezca la  **Cuota de almacenamiento** en el valor más bajo posible (actualmente 110). Así, solo instalará paquetes de complemento en esta colección de sitios, y son muy pequeños.
    
8. Establezca la  **Cuota de recursos de servidor** en 0 (cero). (La cuota de recursos de servidor está relacionada con la limitación de soluciones de espacio aislado de bajo rendimiento, pero no instalarán soluciones de espacio aisladas en el sitio del catálogo de complementos).
    
9. Elija  **Aceptar**.
    
Para agregar un complemento al sitio del catálogo de complementos, vaya al sitio que acaba de crear. En el panel de navegación de la izquierda, elija  **Complementos de Office** y, para cargar un archivo de manifiesto de complemento de Office, elija **nuevo complemento**.    

## <a name="publish-to-an-add-in-catalog"></a>Publicar en un catálogo de complementos


1. Vaya al catálogo de complementos:

    1- Abra la página principal de la Administración central de SharePoint.
    
    2- Seleccione **Complementos**.
    
    3- Seleccione **Administrar catálogo de complementos**.
    
    4- Haga clic en el vínculo que se facilita y elija **Complementos para Office** en la barra de navegación de la izquierda.
    
2. Haga clic en el vínculo **Haga clic para agregar un elemento nuevo**.
    
3. Elija **Examinar** y, después, especifique el [manifiesto](../../docs/overview/add-in-manifests.md) que quiera cargar.
    
    El contenido y los complementos de panel de tareas en este catálogo ahora están disponibles en el cuadro de diálogo **Complementos de Office**. Para tener acceso a ellos, elija **Mis complementos** en la pestaña **Insertar** y, después, elija **MI ORGANIZACIÓN**.
    
Después de cargar el manifiesto del complemento al catálogo de complementos de Office, los usuarios pueden tener acceso a los complementos con el procedimiento siguiente:


1. En la aplicación de Office, vaya a **Archivo**  >  **Opciones**  >  **Centro de confianza**  >  **Configuración del Centro de confianza**  >  **Catálogos de complementos de confianza**.
    
2. Especifique la dirección URL de la _colección de sitios de SharePoint primaria_ del catálogo de complementos. Por ejemplo, si la dirección URL del catálogo de complementos de Office es:
    
    `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    Especifique la dirección URL de la colección de sitios primaria:
    
    `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. Cierre y vuelva a abrir la aplicación de Office. El catálogo de complementos estará disponible en el cuadro de diálogo **Complementos de Office**.
    
Como alternativa, un administrador puede especificar un catálogo de complementos de Office en SharePoint con una directiva de grupo. Para más información, vea la sección "Usar una directiva de grupo para administrar el modo en que los usuarios pueden instalar y usar complementos de Office" en [Información general sobre complementos de Office en TechNet](https://technet.microsoft.com/en-us/library/jj219429.aspx).

