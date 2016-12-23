
# <a name="sideload-office-add-ins-for-testing"></a>Transferir localmente complementos de Office para pruebas

Puede instalar un complemento de Office para realizar pruebas en un cliente de Office que se ejecuta en Windows mediante un catálogo de carpeta compartida para publicar el manifiesto en un recurso compartido de archivos de red. 

>**Nota:** Para probar un complemento de Office en Office Online, consulte [Transferir localmente complementos de Office en Office Online para pruebas](sideload-office-add-ins-for-testing.md). Para probar un complemento en IPad o Mac, consulte [Transferir localmente complementos de Office para pruebas en dispositivos iPad y Mac](sideload-an-office-add-in-on-ipad-and-mac.md ). Para probar un complemento de Outlook, consulte [Transferir localmente complementos de Outlook para pruebas](sideload-outlook-add-ins-for-testing.md ).

Implemente solo el archivo de manifiesto en el catálogo de carpeta compartida. Implemente la aplicación web en un servidor web y especifique la dirección URL en el elemento **SourceLocation** del archivo de manifiesto.

 >**Importante:** Para mejorar la seguridad de las aplicaciones que obtienen acceso a datos y servicios externos, el complemento tiene que usar un protocolo seguro como HTTPS (Hypertext Transfer Protocol Secure) para conectarse a servicios y datos externos. Es necesario usar HTTPS si el complemento usa comandos de complemento.

En el siguiente vídeo, se le guiará por el proceso de transferir de forma local el complemento en Office Online o en la aplicación de escritorio de Office.

<iframe width="560" height="315" src="https://www.youtube.com/embed/XXsAw2UUiQo" frameborder="0" allowfullscreen></iframe>


## <a name="share-a-folder"></a>Compartir una carpeta

1. En el equipo de Windows en el que desee alojar el complemento, vaya a la carpeta principal o a la letra de unidad de la carpeta que quiera usar como catálogo de carpeta compartida.

2. Abra el menú contextual de la carpeta (haga clic con el botón derecho) y elija **Propiedades**.

3. Abra la pestaña **Uso compartido**.

4. En la página **Elegir contactos...**, añádase a usted mismo y añada a las personas con quienes quiera compartir el complemento. Si todos son miembros de un grupo de seguridad, puede añadir el grupo. Se necesita tener al menos permiso de **lectura/escritura** en la carpeta. 

5. Elija **Compartir** > **Hecho** > **Cerrar**.

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>Especificar la carpeta compartida como catálogo de confianza

      
3. Abra un nuevo documento en Excel, Word o PowerPoint.
    
4. Seleccione la pestaña **Archivo** y haga clic en **Opciones**.
    
5. Haga clic en **Centro de confianza** y seleccione el botón **Configuración del Centro de confianza**.
    
6. Elija **Catálogos de complementos de confianza**.
    
7. En el cuadro **URL del catálogo**, escriba la ruta de acceso completa al catálogo de carpetas compartidas y luego elija **Agregar catálogo**.
    
8. Active la casilla **Mostrar en el menú** y haga clic en **Aceptar**.

9. Cierre la aplicación de Office para que los cambios surtan efecto.
    
## <a name="sideload-your-add-in"></a>Transferir localmente el complemento


1. Coloque el archivo de manifiesto de cualquier complemento que se esté probando en el catálogo de carpetas compartidas.

2. En Excel, Word o PowerPoint, seleccione **Mis complementos** en la pestaña **Insertar** de la cinta de opciones.

3. Elija **CARPETA COMPARTIDA** en la parte superior del cuadro de diálogo **Complementos de Office**.

4. Seleccione el nombre del complemento y haga clic en **Aceptar** para insertarlo.


## <a name="additional-resources"></a>Recursos adicionales

- [Validar y solucionar problemas con el manifiesto](troubleshoot-manifest.md)
- [Publicar el complemento de Office](../publish/publish.md)
    
