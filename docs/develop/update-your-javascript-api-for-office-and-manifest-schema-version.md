
# Actualizar la versión de la API de JavaScript para Office y los archivos de esquema de manifiesto



En este artículo se describe cómo actualizar los archivos JavaScript (Office.js y los archivos .js específicos de la aplicación) y el archivo de validación del manifiesto del complemento en el proyecto de complemento de Office a la versión 1.1.

## Usar los archivos de proyecto más actualizados

Si usa Visual Studio para desarrollar su complemento, para usar los [miembros de la API más recientes](../../reference/what's-changed-in-the-javascript-api-for-office.md) de la API de JavaScript para Office y las [características de la versión 1.1 del manifiesto del complemento](../../docs/overview/add-in-manifests.md) (validado frente a offappmanifest-1.1.xsd), tiene que descargar e instalar [Visual Studio 2015 y la versión más reciente de Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs).

Si usa un editor de texto o IDE distinto de Visual Studio para desarrollar su complemento, deberá actualizar las referencias a la red CDN para Office.js y la versión del esquema al que se hace referencia en el manifiesto de la complemento.

Para ejecutar un complemento desarrollado mediante las características nuevas y actualizadas de la API de Office.js y del manifiesto del complemento, los clientes deben ejecutar productos locales de Office 2013 SP1 o una versión posterior y, donde sea aplicable, SharePoint Server 2013 SP1 y productos del servidor relacionados, Exchange Server 2013 Service Pack 1 (SP1) o los productos hospedados en línea equivalentes: Office 365, SharePoint Online y Exchange Online.

Para descargar productos de Office, SharePoint y Exchange SP1, consulte los siguientes temas:


- [Lista de todas las actualizaciones del Service Pack 1 (SP1) de Microsoft Office 2013 y productos de escritorio relacionados](http://support.microsoft.com/kb/2850036)
    
- [Lista de todas las actualizaciones del Service Pack 1 (SP1) de Microsoft SharePoint Server 2013 y productos de servidor relacionados](http://support.microsoft.com/kb/2850035)
    
- [Descripción de Exchange Server 2013 Service Pack 1](http://support.microsoft.com/kb/2926248)
    

## Actualizar un proyecto de Complemento de Office creado con Visual Studio para que use la última biblioteca de la API de JavaScript para Office y la versión 1.1 del esquema del manifiesto del complemento


En el caso de los proyectos que se han creado antes del lanzamiento de la versión 1.1 de la API de JavaScript para Office y el esquema del manifiesto de la aplicación, pueden actualizarse los archivos correspondientes con el  **Administrador de paquetes NuGet**. A continuación, deberán actualizarse las páginas HTML del complemento para que hagan referencia a estos archivos. 

Tenga en cuenta que el proceso de actualización se aplica  _por proyecto_. Deberá repetir el proceso de actualización para cada proyecto de complemento en el que desee usar la versión 1.1 de Office.js y el esquema del manifiesto de la aplicación.




### Para actualizar los archivos de la biblioteca de la API de JavaScript para Office de su proyecto a la versión más reciente:


1. En Visual Studio 2015, abra o cree un nuevo proyecto de  **Complemento de Office**.
    
      - En el panel izquierdo, haga clic en **Actualizar** y complete el proceso de actualización del paquete.
    
  - Vaya al paso 6.
    
2. Elija  **Herramientas**  >  **Administrador de paquetes NuGet**  >  **Administrar paquetes NuGet para la solución**.
    
3. En el  **Administrador de paquetes NuGet**, seleccione  **nuget.org** para **Origen del paquete** y **Actualización disponible** para **Filtro**, y seleccione Microsoft.Office.js.
    
4. En el panel izquierdo, haga clic en **Actualizar** y complete el proceso de actualización del paquete.
    
5. En la etiqueta **head** de las páginas HTML de su complemento, marque como comentario o elimine las referencias existentes al script office.js (por ejemplo, `<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`) y haga referencia a la biblioteca actualizada de la API de JavaScript para Office de la forma siguiente (cambie el valor de versión a "1"). 

   >**Nota** El "/1/" delante de office.js en la dirección URL de CDN especifica el uso de la versión incremental más reciente en la versión 1 de Office.js.
    
```
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```


### Para actualizar el archivo del manifiesto en el proyecto para usar la versión 1.1 del esquema


- En el archivo del manifiesto del complemento del proyecto (_NombreDelProyecto_ Manifest.xml), actualice el atributo **xmlns** del elemento **OfficeApp** y cambie el valor de la versión a "1.1" (no modifique otros atributos que no sean **xmlns**).
    
```XML
  <OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
```


>
  **Nota** Después de actualizar la versión del esquema de manifiesto de complemento a 1.1, tendrá que quitar los elementos **Capabilities** y **Capability** y, después, reemplazarlos por los [elementos Hosts y Host](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx) o por los [elementos Requirements y Requirement](../../docs/overview/specify-office-hosts-and-api-requirements.md).

## Actualizar un proyecto de Complemento de Office creado con un editor de texto u otro IDE para que use la versión más reciente de la biblioteca de la API de JavaScript para Office y del esquema del manifiesto del complemento


Para los proyectos creados antes del lanzamiento de la versión 1.1 de la API de JavaScript para Office y del esquema del manifiesto del complemento, tiene que actualizar las páginas HTML del complemento para que hagan referencia a la red CDN de la biblioteca v1.1 y actualizar el archivo del manifiesto del complemento para que use el esquema v1.1. 

El proceso de actualización se aplica _por proyecto_. Tendrá que repetir el proceso de actualización para cada proyecto de aplicación en el que quiera usar la versión 1.1 de Office.js y del esquema de manifiesto de complemento.

No necesita tener copias locales de los archivos de la API de JavaScript para Office (Office.js y los archivos .js específicos de la aplicación) para desarrollar un complemento de Office (al hacer referencia a Office.js en la CDN, se descargan los archivos necesarios en tiempo de ejecución). Pero, si quiere tener una copia local de los archivos de la biblioteca, puede usar la [utilidad de línea de comandos NuGet](http://docs.nuget.org/consume/installing-nuget) y el comando `Install-Package Microsoft.Office.js` para descargarlos.

 > **Nota** Para obtener una copia del XSD (definición del esquema XML) para el manifiesto del complemento versión 1.1, vea el listado en [Schema reference for Office Add-ins manifests (v1.1)](../overview/add-in-manifests.md) [Referencia de esquema de manifiestos de complementos de Office (versión 1.1)].


### Para actualizar los archivos de la biblioteca de la API de JavaScript para Office de su proyecto a la versión más reciente


1. Abra las páginas HTML del complemento en su editor de texto o IDE.
    
2. En la etiqueta **head** de las páginas HTML de su complemento, marque como comentario o elimine las referencias existentes al script office.js (por ejemplo, `<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`) y haga referencia a la biblioteca actualizada de la API de JavaScript para Office de la forma siguiente (cambie el valor de versión a "1").
    
```
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```


    The  `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.
    

### Para actualizar el archivo del manifiesto en el proyecto para usar la versión 1.1 del esquema


- En el archivo del manifiesto del complemento del proyecto ( _nombre_de_proyecto_ Manifest.xml), actualice el atributo **xmlns** del elemento **OfficeApp** y cambie el valor de la versión a `1.1` (no modifique otros atributos que no sean **xmlns**).
    
```XML
<OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
```

>
  **Nota** Después de actualizar la versión del esquema de manifiesto de complemento a 1.1, tendrá que quitar los elementos **Capabilities** y **Capability** y, después, reemplazarlos por los [elementos Hosts y Host](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx) o por los [elementos Requirements y Requirement](../../docs/overview/specify-office-hosts-and-api-requirements.md).
    

## Recursos adicionales



- [Especificar los hosts de Office y los requisitos de la API](../../docs/overview/specify-office-hosts-and-api-requirements.md)
    
- [Información sobre la API de JavaScript para Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [API de JavaScript para Office](../../reference/javascript-api-for-office.md)
    
- [Referencia de esquema para manifiestos de Complementos de Office (v1.1)](../overview/add-in-manifests.md)
    
