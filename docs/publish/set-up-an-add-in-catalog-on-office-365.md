
# Configurar un catálogo de complementos en Office 365

Un catálogo de complemento es una colección de sitios dedicados de una aplicación web de SharePoint o un inquilino de SharePoint Online que hospeda bibliotecas de documentos para Complementos de SharePoint y Complementos de Office. Para poder usar los archivos de manifiesto de las Complementos de Office en las organizaciones, los administradores pueden cargarlos en el catálogo de complementos. Si un administrador registra un catálogo de complementos como catálogo de confianza (estableciendo una directiva de grupo o especificando el catálogo de confianza en la pestaña  **Catálogos de complementos de confianza** del cuadro de diálogo **Opciones** por medio de **Archivo** > **Opciones** > **Centro de confianza** > **Configuración del Centro de confianza** > **Catálogos de complementos de confianza**), los usuarios podrán insertar el complemento desde la interfaz de usuario de inserción de una aplicación cliente de Office.

## Para configurar un catálogo de complemento en SharePoint Online


1. En la página del centro de administración de Office 365, elija  **Administrador** y luego **SharePoint**.
    
2. En el panel de tareas de la izquierda, elija  **complementos**.
    
3. En la página  **complementos**, elija  **Catálogo de complementos**.
    
4. En la página  **Sitio del catálogo de complementos**, elija  **Aceptar** para aceptar la opción predeterminada y crear un nuevo sitio del catálogo de complementos.
    
5. En la página  **Crear colección de sitios del catálogo de complementos**, especifique el título del sitio del catálogo de complementos.
    
6. Especifique la dirección del sitio web.
    
7. Establezca la  **Cuota de almacenamiento** en el valor más bajo posible (actualmente 110). Así, solo instalará paquetes de complemento en esta colección de sitios, y son muy pequeños.
    
8. Establezca la  **Cuota de recursos de servidor** en 0 (cero). (La cuota de recursos de servidor está relacionada con la limitación de soluciones de espacio aislado de bajo rendimiento, pero no instalarán soluciones de espacio aisladas en el sitio del catálogo de complementos).
    
9. Elija  **Aceptar**.
    
Para agregar un complemento al sitio del catálogo de complementos, vaya al sitio que acaba de crear. En el panel de navegación de la izquierda, elija  **Complementos de Office** y, para cargar un archivo de manifiesto de complemento de Office, elija **nuevo complemento**.


## Recursos adicionales


- [Publicar complementos en un catálogo de complementos](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)

    

