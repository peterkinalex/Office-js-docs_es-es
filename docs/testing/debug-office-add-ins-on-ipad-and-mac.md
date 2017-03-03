# <a name="debug-office-add-ins-on-ipad-and-mac"></a>Depurar complementos de Office en dispositivos iPad y Mac

Puede usar Visual Studio para desarrollar y depurar add-ins en Windows, pero no se puede usar para depurar add-ins en un iPad ni en un Mac. Como los add-ins se desarrollan en HTML y Javascript, están diseñados para que funcionen en distintas plataformas, pero puede haber diferencias sutiles en la forma en que cada explorador presenta el código HTML. En este tema se describe cómo depurar los add-ins que se ejecutan en iPad o Mac. 

## <a name="debugging-with-vorlonjs"></a>Depurar con Vorlon.JS 

Vorlon.JS es un depurador de páginas web, similar a las herramientas F12. Está diseñado para trabajar de manera remota y le permite depurar páginas web en diferentes dispositivos. Para obtener más información, vea el [sitio web de Vorlon](http://www.vorlonjs.com).  

### <a name="install-and-set-up-up-vorlonjs-on-a-mac-or-ipad"></a>Instalar y configurar Vorlon.JS en un dispositivo Mac o iPad 

1.    Inicie sesión en el dispositivo como administrador.

2.    Instale [Node.js](https://nodejs.org) si todavía no está instalado. 

2.    Abra una ventana **Terminal** y escriba el comando `npm i -g vorlon`. La herramienta está instalada en `/usr/local/lib/node_modules/vorlon`.

### <a name="configure-vorlonjs-to-use-https"></a>Configurar Vorlon.JS para usar HTTPS

Para depurar una aplicación que usa Vorlon.JS, agregue una etiqueta `<script>` a la página inicial de la aplicación que carga un script de Vorlon.JS desde una ubicación conocida (para obtener información, vea el siguiente procedimiento). Los complementos necesitan el protocolo HTTPS, es decir, SSL. Por extensión, cualquier script que usen debe hospedarse desde un servidor HTTPS, incluido el script de Vorlon.JS. Por lo tanto, tiene que configurar Vorlon.JS para usar SSL y así poder usar Vorlon.JS con complementos. 

4.    En **Buscador**, vaya a `/usr/local/lib/node_modules/vorlon`, abra el menú contextual (botón derecho) para la carpeta `/Server` y, después, seleccione **Obtener información**.

5.    Pulse el icono de candado en la esquina inferior derecha de la ventana **Información del servidor** para desbloquear la carpeta.

6. En la sección **Uso compartido y permisos** de la ventana, establezca el **Privilegio** para el grupo **Personal** en **Lectura y escritura**.

7. Pulse el icono de candado de nuevo para ***volver a bloquear*** la carpeta.

8. De nuevo en **Buscador**, expanda la subcarpeta `/Server`, haga clic con el botón derecho en el archivo `config.json` y, después, seleccione **Obtener información**.

9. En la ventana **Información de config.json**, cambie los privilegios del archivo exactamente de la manera que lo hizo para su carpeta `/Server` primaria. Asegúrese de volver a bloquear y cerrar la ventana.

10. De nuevo en **Buscador**, haga clic con el botón derecho en el archivo `config.json`, seleccione **Abrir con** y, después, seleccione **TextEdit**. El archivo se abre en un editor de texto.

11. Cambie el valor de la propiedad **useSSL** a `true`.

12. En la sección **Complementos**, busque el complemento con el **id.** de `OFFICE` y el **nombre** de `Office Addin`. Si la propiedad **habilitada** del complemento ya no es `true`, establézcala en `true`.

13. Guarde el archivo y cierre el editor.

5.    En **Buscador**, vaya a `/usr/local/lib/node_modules/vorlon`, haga clic con el botón derecho en la subcarpeta `Server` y seleccione **Nuevo terminal en la carpeta**. 
    
7.    En la ventana **Terminal**, escriba `sudo vorlon`. Se le pedirá que escriba su contraseña de administrador. El servidor Vorlon se inicia. Deje la ventana **Terminal** abierta.

6.    Abra una ventana del explorador y vaya a `https://localhost:1337`, que es la interfaz de Vorlon.JS. Cuando se le pida, pulse **Siempre** para confiar en el certificado de seguridad. 

    >**Nota:** Si no se le pide, puede que necesite confiar en el certificado de manera manual. El archivo de certificado es `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Pruebe los siguientes pasos. Si tiene problemas, vea la ayuda de Macintosh o iPad. 
    >
    >1. Cierre la ventana del explorador y, en la ventana **Terminal** que está ejecutando el servidor Vorlon, use CTRL-C para detenerlo.
    >2. En **Buscador**, haga clic con el botón derecho en el archivo `server.crt` y seleccione **Acceso a llaves**. La ventana **Acceso a llaves** se abre.
    >2. En la lista **Llaves** de la izquierda, seleccione **Iniciar sesión** si ya no está seleccionado y, después, seleccione **Certificados** en la sección **Categoría**. Se muestra el certificado **localhost**.
    >3. Haga clic con el botón derecho en el certificado **localhost** y seleccione **Obtener información**. Se abre una ventana **localhost**.
    >4. En la sección **Confianza**, abra el selector **Al usar este certificado** y seleccione **Confiar siempre**. 
    >5. Cierre la ventana **localhost**. Si la acción se ha realizado correctamente, el certificado **localhost** de la ventana **Acceso a llaves** tiene una cruz blanca en un círculo azul en su icono.

### <a name="configure-the-add-in-for-vorlonjs-debugging"></a>Configurar el complemento para la depuración de Vorlon.JS

1. Agregue la siguiente etiqueta de script a la sección `<head>` del archivo home.html (o el archivo HTML principal) del complemento:
```    
<script src="https://localhost:1337/vorlon.js"></script>    
```  

2. Implemente la aplicación web del complemento en un servidor web que sea accesible desde el dispositivo Mac o iPad, como un sitio web de Azure. 

3. Actualice la dirección URL del complemento en todos los lugares donde la dirección URL aparece en el manifiesto del complemento.

4. Copie el manifiesto del complemento en la siguiente carpeta del dispositivo Mac o iPad: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, donde *{host_name}* es Word, Excel, PowerPoint u Outlook.

### <a name="inspect-an-add-in-in-vorlonjs"></a>Inspeccionar un complemento en Vorlon.JS

1. Si el servidor Vorlon no se está ejecutando, en **Buscador**, vaya a `/usr/local/lib/node_modules/vorlon`, haga clic con el botón derecho en la subcarpeta `Server` y seleccione **Nuevo terminal en la carpeta**. 
    
7.    En la ventana **Terminal**, escriba `sudo vorlon`. Se le pedirá que escriba su contraseña de administrador. El servidor Vorlon se inicia. Deje la ventana **Terminal** abierta.

6.    Abra una ventana del explorador y vaya a `https://localhost:1337`, que es la interfaz de Vorlon.JS.

7. Transfiera el complemento localmente. Si es para Excel, PowerPoint o Word, transfiéralo localmente como se describe en [Sideload an Office Add-in on iPad and Mac](https://dev.office.com/docs/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac) (Transferir un complemento de Office localmente en un dispositivo iPad y Mac). Si es un complemento de Outlook, transfiéralo localmente como se describe en [Sideload Outlook Add-ins for testing](https://dev.office.com/docs/add-ins/testing/sideload-outlook-add-ins-for-testing) (Transferir localmente complementos de Outlook para pruebas). Si el complemento no usa comandos de complemento, se abrirá inmediatamente. De otro modo, pulse el botón para abrir el complemento. Dependiendo de la compilación de la aplicación host de Office, el botón estará en la pestaña **Inicio** o en una pestaña **Complemento**.

El complemento se mostrará en la lista de clientes de Vorlon.JS (en el lateral izquierdo de la interfaz de Vorlon.JS) como **{OS} - n**, para algún número *n* y donde *{OS}* es el tipo de dispositivo, como "Macintosh". 

![Captura de pantalla donde se muestra la interfaz de Vorlon.js](../../images/vorlon_interface.png)

La herramienta de Vorlon tiene una variedad de complementos. Los que están habilitados en estos momentos aparecen como pestañas en la parte superior de la herramienta. (Puede habilitar más complementos mediante la elección del icono de engranaje a la izquierda). Estos complementos son similares a las funciones de las herramientas F12. Por ejemplo, pueden resaltar elementos DOM, ejecutar comandos y mucho más. Para obtener más información, vea [Vorlon Documentation Core Plugins](http://vorlonjs.com/documentation/#console) (Complementos principales de la documentación de Vorlon). 

Un **complemento de Office** agrega características adicionales a Office.js, como la exploración del modelo de objetos, la ejecución de llamadas de Office.js y la lectura de los valores de las propiedades de objeto. Para obtener instrucciones, vea [VorlonJS plugin for debugging Office Addin](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/) (Complemento VorlonJS para depurar complementos de Office).

>**Nota:** No hay ninguna manera de establecer puntos de interrupción en Vorlon.JS.

## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a>Borrar la memoria caché de la aplicación Office en un dispositivo Mac o iPad

A menudo en Office para Mac los complementos se almacenan en caché, por motivos de rendimiento. Normalmente, la memoria caché se borra volviendo a cargar el complemento. Si existe más de un complemento en el mismo documento, el proceso de borrar la caché automáticamente durante la recarga puede no ser fiable. 

En un dispositivo Mac, puede borrar la memoria caché manualmente eliminando todo de la carpeta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`. 

En un dispositivo iPad, puede llamar a `window.location.reload(true)` desde JavaScript en el complemento para forzar una recarga. De manera alternativa, puede volver a instalar Office.
