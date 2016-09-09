
# Depurar complementos en Office Online


Se pueden compilar y depurar complementos en un equipo que no ejecute Windows o con el cliente de escritorio de Office 2013 u Office 2016 (por ejemplo, si desarrolla en un equipo Mac). En este artículo se describe cómo usar Office Online para probar y depurar complementos. 

Para empezar:


- Obtenga una cuenta de desarrollador de Office 365, si aún no la tiene, u obtenga acceso a un sitio de SharePoint.
    
     >**Nota**  Para obtener una cuenta gratuita de desarrolladores de Office 365, únase a nuestro [programa de desarrolladores de Office 365](https://dev.office.com/devprogram).
     
- Configure un catálogo de complementos en Office 365 (SharePoint Online). Un catálogo de complementos es una colección de sitios dedicada en SharePoint Online que contiene bibliotecas de documentos para los complementos de Office. Si dispone de su propio sitio de SharePoint, puede configurar una biblioteca de documentos de catálogo de complementos. Para obtener más información, consulte [Publicar complementos de panel de tareas y de contenido en un catálogo de complementos de SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).
    

## Depurar el complemento desde Excel Online o Word Online

Para depurar el complemento mediante Office Online:


1. Implemente el complemento en un servidor compatible con SSL.
    
     >**Nota:**  Le recomendamos que use el [Generador Yeoman](https://github.com/OfficeDev/generator-office) para crear y alojar el complemento.
     
2. En el [archivo de manifiesto del complemento](../../docs/overview/add-in-manifests.md), actualice el valor del elemento  **SourceLocation** de modo que incluya un URI absoluto en lugar de relativo. Por ejemplo:
    
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. Cargue el manifiesto en la biblioteca de complementos de Office del catálogo de complementos en SharePoint.
    
4. Inicie Excel Online o Word Online desde el iniciador de aplicaciones de Office 365 y abra un nuevo documento.
    
5. En la pestaña Insertar, seleccione  **Mis complementos** o **Complementos de Office** para insertar el complemento y probarlo en la aplicación.
    
6. Use el depurador de herramientas del explorador que prefiera para depurar el complemento.
    
    Los siguientes son algunos problemas que pueden surgir durante la depuración:
    
  - Algunos de los errores de JavaScript que experimente pueden originarse en Office Online.
    
  - El explorador puede mostrar un error de certificado no válido que tendrá que pasar por alto.
    
  - Si establece puntos de interrupción en el código, Office Online puede producir un error que indica que no se puede guardar.
    

## Recursos adicionales


- [Procedimientos recomendados para desarrollar complementos para Office](../overview/add-in-development-best-practices.md)
    
- [Directivas de validación para aplicaciones enviadas a la Tienda Office (versión 1.9)](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)
    
- [Crear complementos y aplicaciones de la Tienda de Office efectivos](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)
    
- [Solucionar errores de usuario con los complementos de Office](../testing/testing-and-troubleshooting.md)
    
