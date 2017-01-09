
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>Referencia a la biblioteca de la API de JavaScript para Office desde su red de entrega de contenido (CDN)


La biblioteca de la [API de JavaScript para Office](../../reference/javascript-api-for-office.md) consta del archivo Office.js y los archivos .js específicos de la aplicación host asociada, como Excel-15.js y Outlook-15.js. 


La forma más sencilla de hacer referencia a la API consiste en usar la red CDN añadiendo el siguiente `<script>` a la etiqueta `<head>` de la página:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

El  `/1/` delante de `office.js` en la dirección URL de CDN especifica la versión incremental más reciente dentro de la versión 1 de Office.js. Debido a que la API de JavaScript para Office mantiene la compatibilidad con versiones anteriores, la versión más reciente seguirá admitiendo miembros de la API que se incorporaron anteriormente en la versión 1. Si necesita actualizar un proyecto existente, consulte [Actualizar la versión de la API de JavaScript para Office y los archivos de esquema de manifiesto](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Si tiene previsto publicar su complemento de Office desde la Tienda Office, tendrá que usar esta referencia de red CDN. Las referencias locales solo son adecuadas para escenarios internos, de desarrollo y de depuración.

> **Importante:** Al desarrollar un complemento para una aplicación host de Office, haga referencia a la API de JavaScript para Office desde una sección `<head>` de la página. Esto garantiza que la API se inicializa por completo antes de los elementos del cuerpo. Los hosts de Office necesitan que los complementos se inicialicen 5 segundos después de la activación. Si el complemento no se activa dentro de este umbral, se declarará como que no responde y se mostrará un mensaje de error al usuario.       

## <a name="additional-resources"></a>Recursos adicionales



- [Información sobre la API de JavaScript para Office](../../docs/develop/understanding-the-javascript-api-for-office.md)    
- [API de JavaScript para Office](../../reference/javascript-api-for-office.md)
    
