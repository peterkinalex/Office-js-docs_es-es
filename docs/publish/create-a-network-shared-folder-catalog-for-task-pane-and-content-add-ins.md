
# Crear un catálogo de carpetas compartidas de red para complementos de panel de tareas y de contenido


Un catálogo de carpeta compartida permite publicar los manifiestos de los Complementos de Office de contenido y del panel de tareas en un recurso compartido de archivos de red. Para obtener complementos, los usuarios deben establecer este recurso compartido de archivos como un catálogo de confianza con el procedimiento siguiente.

El archivo de manifiesto es un archivo XML que permite describir mediante declaración cómo debe activarse el complemento cuando el usuario final lo instale y use con las aplicaciones y los documentos de Office. Para más información, consulte [Manifiesto XML de complementos para Office](../../docs/overview/add-in-manifests.md).

En el catálogo de carpetas compartida solo debe implementarse el archivo de manifiesto. La aplicación web propiamente dicha debe implementarse en el servidor web y especificarse su dirección URL en el elemento  **SourceLocation** del archivo de manifiesto.

 >**Importante:** Para mejorar la seguridad de las aplicaciones que obtienen acceso a datos y servicios externos, el complemento tiene que usar un protocolo seguro como HTTPS (Hypertext Transfer Protocol Secure) para conectarse a servicios y datos externos. Es necesario usar HTTPS si el complemento usa comandos de complemento.


## Especificar un recurso compartido de archivos como catálogo de confianza


1. Cree una carpeta en un recurso compartido de red; por ejemplo:  `\\MyShare\MyManifests`.
    
2. Coloque los archivos del manifiesto para los complementos de contenido y panel de tareas que quiere publicar en este recurso compartido de archivos.
    
3. Abra un nuevo documento en Excel, Word o PowerPoint.
    
4. Elija la pestaña  **Archivo** y luego **Opciones**.
    
5. Elija  **Centro de confianza** y elija el botón **Configuración del Centro de confianza**.
    
6. Elija  **Catálogos de complementos de confianza**.
    
7. En el cuadro  **URL de catálogo**, escriba la ruta de acceso al recurso compartido de red que creó en el paso 1 y luego elija  **Agregar catálogo**.
    
8. Active la casilla  **Mostrar en menú** y luego elija **Aceptar**.
    
Tras completar estos pasos, si desea insertar un complemento de contenido o de panel de tareas desde este catálogo, seleccione  **Mis complementos** en la pestaña **Insertar** de la cinta de opciones y elija **Carpeta compartida** en la parte superior del cuadro de diálogo **Complementos de Office**.

Cualquier archivo de manifiesto adicional que agregue a este recurso compartido de archivos estará disponible para los usuarios que hayan especificado esta carpeta compartida.


## Recursos adicionales



- [Publicar el complemento para Office](../publish/publish.md)
    

