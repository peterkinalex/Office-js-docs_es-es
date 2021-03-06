# <a name="build-your-first-onenote-add-in"></a>Crear el primer complemento de OneNote

En este artículo se indican los pasos necesarios para crear un complemento de panel de tareas sencillo que agrega texto a una página de OneNote.

En la imagen siguiente se muestra el complemento que creará.

   ![El complemento de OneNote que se crea con este tutorial](../../images/onenote-first-add-in.png)

<a name="setup"></a>
## <a name="step-1-set-up-your-dev-environment-and-create-an-add-in-project"></a>Paso 1: Configurar el entorno de desarrollo y crear un proyecto de complemento
Siga las instrucciones para [Crear un complemento de Office con cualquier editor](../get-started/create-an-office-add-in-using-any-editor.md) para instalar los requisitos previos necesarios y ejecutar el generador de Yeoman Office para crear un nuevo proyecto de complemento. En la tabla siguiente se enumeran los atributos del proyecto para seleccionar en el generador de Yeoman.

| Opción | Valor |
|:------|:------|
| Subcarpeta nueva: | (acepte los valores predeterminados) |
| Nombre de complemento | Complemento de OneNote |
| Aplicación de Office compatible | (seleccione OneNote) |
| Crear complemento | Sí, quiero un complemento nuevo |
| Agregar [TypeScript](https://www.typescriptlang.org/) | No |
| Elegir marco | Jquery |

<a name="develop"></a>
## <a name="step-2-modify-the-add-in"></a>Paso 2: Modificar el complemento
Puede editar el complemento con cualquier editor de texto o IDE. Si aún no ha probado Visual Studio Code, puede [descargarlo de forma gratuita](https://code.visualstudio.com/) en Linux, Mac OS X y Windows.

1. Abra **index.html** en el directorio del proyecto. 

2. Reemplace el elemento `<main>` con el código siguiente. Esto agrega un área de texto y un botón mediante [componentes de Office UI Fabric](http://dev.office.com/fabric/components).

```html
<main class="ms-welcome__main">
   <br />
   <p class="ms-font-l">Enter content below</p>
   <div class="ms-TextField ms-TextField--placeholder">
       <textarea id="textBox" rows="5"></textarea>
   </div>
   <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
        <span class="ms-Button-label">Add Outline</span>
        <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
        <span class="ms-Button-description">Adds the content above to the current page.</span>
    </button>
</main>
```

3. Abra **app.js** (o app.ts si usa TypeScript) en el directorio del proyecto. Edite la función **Office.initialize** para agregar un evento de clic al botón **Add outline** como se indica a continuación.

```js
// The initialize function is run each time the page is loaded.
Office.initialize = function (reason) {
   $(document).ready(function () {
       app.initialize();
       
       // Set up event handler for the UI.
       $('#addOutline').click(addOutlineToPage);
   });
};
```
 
4. Reemplace el método **un** con el método siguiente **addOutlineToPage**. Esto obtiene el contenido del área de texto y lo agrega a la página.

```js
// Add the contents of the text area to the page.
function addOutlineToPage() {        
   OneNote.run(function (context) {
      var html = '<p>' + $('#textBox').val() + '</p>';
      
       // Get the current page.
       var page = context.application.getActivePage();
       
       // Queue a command to load the page with the title property.             
       page.load('title'); 
       
       // Add an outline with the specified HTML to the page.
       var outline = page.addOutline(40, 90, html);
       
       // Run the queued commands, and return a promise to indicate task completion.
       return context.sync()
           .then(function() {
               console.log('Added outline to page ' + page.title);
           })
           .catch(function(error) {
               app.showNotification("Error: " + error); 
               console.log("Error: " + error); 
               if (error instanceof OfficeExtension.Error) { 
                   console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
               } 
           }); 
       });
}
```

<a name="test"></a>
## <a name="step-3-test-the-add-in-on-onenote-online"></a>Paso 3: Probar el complemento en OneNote Online
1. Inicie el servidor HTTPS.  

  a. Abra un símbolo del sistema o una instancia de Terminal **cmd** y vaya a la carpeta del proyecto del complemento. 
  
  b. Ejecute el comando, como se muestra a continuación.

  ```
  C:\your-local-path\onenote add-in\> npm start
  ```

2. Instale el certificado autofirmado como un certificado de confianza. Solo necesita realizar esto una vez en el equipo para los proyectos de complemento creados con el generador Yeoman Office. Para obtener más información, consulte [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) (Agregar certificados autofirmados como certificado raíz de confianza).

3. Vaya a [OneNote Online](https://www.onenote.com/notebooks) y abra un bloc de notas.

4. Elija **Insertar > Complementos de Office**. Se abrirá el cuadro de diálogo de complementos de Office.

  -Si ha iniciado sesión con su cuenta de consumidor, elija la pestaña **MIS COMPLEMENTOS** y después elija **Cargar mi complemento**.
  
  -Si ha iniciado sesión con su cuenta escolar o de trabajo, elija la pestaña **MI ORGANIZACIÓN** y después elija **Cargar mi complemento**. 
  
  La siguiente imagen muestra la pestaña **MIS COMPLEMENTOS** para blocs de notas de consumidor.

  ![El cuadro de diálogo Complementos de Office con la pestaña MIS COMPLEMENTOS](../../images/onenote-office-add-ins-dialog.png)

5. En el diálogo Cargar complemento, busque **onenote-add-in-manifest.xml** en la carpeta del proyecto y, a continuación, elija **Cargar**. Durante las pruebas, se almacenará el archivo de manifiesto en el almacenamiento local del explorador.

6. El complemento se abre en un iFrame junto a la página de OneNote. Escriba texto en el área de texto y, después, elija **Agregar esquema**. El texto que escriba se agregará a la página. 

## <a name="troubleshooting-and-tips"></a>Solución de problemas y sugerencias
-Puede depurar el complemento con las herramientas de desarrollo del navegador. Si usa el servidor web Gulp y realiza la depuración en Internet Explorer o en Chrome, puede guardar los cambios de forma local y, después, actualizar el iFrame de los complementos.

-Al inspeccionar un objeto de OneNote, las propiedades que están disponibles actualmente para su uso se muestran como valores reales. Las propiedades que necesitan cargarse se muestran como *undefined*. Expanda el nodo `_proto_` para ver las propiedades que se han definido en el objeto pero aún no se han cargado.

![Objeto de OneNote no cargado en el depurador](../../images/onenote-debug.png)

-Necesita habilitar el contenido mixto en el explorador si el complemento usa algún recurso HTTP. Los complementos de producción solo deberían usar recursos HTTPS seguros.

-Los complementos del panel de tareas se pueden abrir desde cualquier lugar, pero solo se pueden insertar complementos de contenido en contenido de páginas normales (es decir, no en títulos, imágenes, iFrames, etc.). 

## <a name="additional-resources"></a>Recursos adicionales

-[Introducción a la programación de API de JavaScript para OneNote](onenote-add-ins-programming-overview.md)

-[Referencia de la API de JavaScript de OneNote](../../reference/onenote/onenote-add-ins-javascript-reference.md)

-[Ejemplo de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)

-[Introducción la plataforma de complementos para Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
