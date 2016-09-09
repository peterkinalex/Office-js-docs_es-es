
# Abordar las limitaciones de la directiva del mismo origen en complementos para Office


La directiva del mismo origen exigida por el explorador impide que un script que se carga desde un dominio pueda obtener o manipular las propiedades de una página web de otro dominio. Esto significa que, de manera predeterminada, el dominio de una dirección URL solicitada debe ser el mismo que el dominio de la página web actual. Por ejemplo, esta directiva impedirá que una página web en un dominio realice llamadas a servicios web [XmlHttpRequest ](http://www.w3.org/TR/XMLHttpRequest/) a otro dominio que no sea el dominio donde se hospeda.

Como los Complementos de Office se hospedan en un control de explorador, la directiva del mismo origen también se aplica a los scripts que se ejecutan en sus páginas web.

Para abordar el cumplimiento de directivas del mismo origen en el desarrollo de los complementos, puede:

- Usar JSON/P para obtener acceso de manera anónima 
    
- Implementar un script del lado del servidor con un esquema de autenticación basado en tokens
    
- Compartir recursos de origen cruzado (CORS)
    
- Crear su propio proxy con IFRAME y POST MESSAGE
    

## Uso de JSON/P para un acceso anónimo


Una forma de superar esta limitación es usar JSON/P para proporcionar un proxy para el servicio web. Para ello, incluya una etiqueta de `script` con un atributo `src` que apunte a algún script hospedado en cualquier dominio. Puede crear mediante programación las etiquetas de `script`, crear dinámicamente la dirección URL para apuntar al atributo `src` y, después, pasar parámetros a la dirección URL a través de parámetros de consulta URI. Los proveedores de servicios web crean y hospedan código JavaScript en direcciones URL específicas y devuelven scripts diferentes dependiendo de los parámetros de consulta URI. Después, estos scripts se ejecutan donde se insertan y funcionan del modo previsto.

Aquí se muestra un ejemplo de JSON/P con una técnica que funcionará en cualquier complemento de Office.

```js
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}

```


## Implementación de un script del lado del servidor con un esquema de autenticación basado en tokens


Otra forma de resolver las limitaciones de la directiva del mismo origen es implementar la página web del complemento como una página ASP que usa OAuth o que almacena en caché las credenciales de las cookies.

Para ver un ejemplo de cómo se usa OAuth para la autenticación, consulte [Twitter SharePoint web part with OAuth](http://aidangarnish.net/post/Twitter-SharePoint-Web-Part-With-OAuth) (Elemento web de Twitter en SharePoint web con OAuth).

Para ver un ejemplo de código del lado del servidor en el que se muestra cómo usar el objeto de `Cookie` en `System.Net` para obtener y establecer valores de las cookies, vea la propiedad [Value](http://msdn2.microsoft.com/EN-US/library/4f772twc).


## Uso compartido de recursos entre orígenes (CORS)


Puede ver un ejemplo de uso de la característica de uso compartido de recursos entre orígenes de [XmlHttpRequest2](http://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) en la sección "Uso compartido de recursos entre orígenes (CORS)" de la página [Nuevos consejos para XMLHttpRequest2](http://www.html5rocks.com/en/tutorials/file/xhr2/).


## Creación de un proxy propio con IFRAME y POST MESSAGE


Para ver un ejemplo de cómo crear su propio proxy con IFRAME y POST MESSAGE, vea el tema sobre [mensajería en varias ventanas](http://ejohn.org/blog/cross-window-messaging/).


## Recursos adicionales


- [Privacidad y seguridad de complementos para Office](../../docs/develop/privacy-and-security.md)
    
