
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-(cdn)"></a>Referencia a la biblioteca de la API de JavaScript para Office desde su red de entrega de contenido (CDN)


La biblioteca de la [API de JavaScript para Office](../../reference/javascript-api-for-office.md) está formada por el archivo Office.js y los archivos .js específicos de la aplicación host asociada, como Excel-15.js y Outlook-15.js. 


El método más sencillo para hacer referencia a la API es usar nuestra CDN. Para hacerlo, agregue el siguiente `<script>` a la etiqueta `<head>` de la página:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

El `/1/` delante de `office.js` en la URL de CDN especifica que es necesario usar la versión incremental más reciente de la versión 1 de Office.js. Como la API de JavaScript para Office mantiene la compatibilidad con versiones anteriores, la versión más reciente seguirá siendo compatible con los miembros de la API que se introdujeron en la versión 1. Si necesita actualizar un proyecto existente, consulte [Actualizar la versión de la API de JavaScript para Office y los archivos de esquema del manifiesto] (../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Si tiene previsto publicar su complemento de Office desde la Tienda Office, tendrá que usar esta referencia de CDN. Las referencias locales solo son adecuadas para escenarios internos, de desarrollo y de depuración.

> **Importante:** Al desarrollar un complemento para una aplicación host de Office, es importante hacer referencia a la API de JavaScript para Office desde dentro de la sección `<head>` de la página. Esto garantiza que la API se inicializa por completo antes que los elementos de body. Los hosts de Office necesitan que los complementos se inicialicen 5 segundos después de la activación. Al superar este umbral, el complemento no responde y se muestra un mensaje de error al usuario.       

## <a name="additional-resources"></a>Recursos adicionales



- [Información sobre la API de JavaScript para Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Office Add-ins platform overview (Información general sobre la plataforma de complementos para Office)](../../docs/overview/office-add-ins.md)
    
- [Ciclo de vida de desarrollo de complementos de Office](../../docs/design/add-in-development-lifecycle.md)
    
- [API de JavaScript para Office](../../reference/javascript-api-for-office.md)
    
