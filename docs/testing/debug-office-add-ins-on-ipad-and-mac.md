
# Depurar complementos de Office en dispositivos iPad y Mac

Puede usar Visual Studio para desarrollar y depurar add-ins en Windows, pero no se puede usar para depurar add-ins en un iPad ni en un Mac. Como los add-ins se desarrollan en HTML y Javascript, están diseñados para que funcionen en distintas plataformas, pero puede haber diferencias sutiles en la forma en que cada explorador presenta el código HTML. En este tema se describe cómo depurar los add-ins que se ejecutan en iPad o Mac. 

## Depurar con Vorlon.js 

Vorlon.js es un depurador de páginas web, similar a las herramientas de F12, que está diseñado para trabajar de forma remota y permite depurar páginas web en diferentes dispositivos. Para más información, vea el [sitio web de Vorlon](http://www.vorlonjs.com).  

Para instalar y configurar Vorlon: 

1.  Instale [Node.js](https://nodejs.org) (si aún no lo hizo). 

2.  Instale Vorlon con npm con el comando siguiente: `sudo npm i -g vorlon` 

3.  Ejecute el servidor Vorlon con el comando `vorlon`. 

4.  Abra una ventana del explorador y vaya a [http://localhost:1337](http://localhost:1337), que es la interfaz de Vorlon.

5.  Agregue la siguiente etiqueta de script a la sección `<head>` del archivo home.html (o el archivo HTML principal) del complemento:
```    
<script src="http://localhost:1337/vorlon.js"></script>    
```  

>**Nota:** Debe habilitar HTTPS en Vorlon para usar Vorlon.js para depurar complementos. Para obtener información sobre cómo hacerlo, consulte [VorlonJS plugin for debugging Office Addin](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/) (Complemento VorlonJS para depurar complementos de Office).

Ahora, cuando abra el complemento en un dispositivo, se mostrará en la lista de clientes en Vorlon (en la parte izquierda de la interfaz de Vorlon). Puede resaltar de forma remota elementos DOM, ejecutar comandos de forma remota y mucho más.  

![Captura de pantalla donde se muestra la interfaz de Vorlon.js](../../images/vorlon_interface.png)

Un complemento de Vorlon dedicado para complementos de Office que agrega funciones adicionales, como interactuar con las API de Office.js. Para más información, vea la entrada de blog [Complemento VorlonJS para depurar complementos de Office](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/). Para habilitar el complemento de complementos de Office: 

1.  Clone de forma local la rama para desarrolladores del repositorio de GitHub Vorlon.js con los comandos siguientes: 
```
git clone https://github.com/MicrosoftDX/Vorlonjs.git
git checkout dev
npm install
```

2.  Abra el archivo **config.json** ubicado en /Vorlon/Server/config.json. Para activar el complemento de complemento de Office, establezca la propiedad **enabled** en **true**.

![Captura de pantalla donde se muestra la sección de complementos de config.json](../../images/vorlon_plugins_config.png) 
