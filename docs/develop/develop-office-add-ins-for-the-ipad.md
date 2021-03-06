
# <a name="develop-office-add-ins-for-the-ipad"></a>Desarrollar complementos de Office para iPad


En la tabla siguiente se muestran las tareas que es necesario realizar para desarrollar un complemento de Office para que se ejecute en Office para iPad.


|**Tarea**|**Descripción**|**Recursos**|
|:-----|:-----|:-----|
|Actualice el complemento para que admita Office.js versión 1.1.|Actualice a la versión 1.1 los archivos de JavaScript (archivos .js específicos de la aplicación y Office.js) y el archivo de validación del manifiesto del complemento usados en el proyecto de complemento de Office.|[Novedades en la API de JavaScript para Office](../../reference/what's-changed-in-the-javascript-api-for-office.md)|
|Aplique los procedimientos recomendados de diseño de interfaz de usuario.|Integre a la perfección la interfaz de usuario del complemento con la experiencia de iOS.|[Diseño para iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Aplique los procedimientos recomendados de diseño de complemento.|Asegúrese de que el complemento proporciona un valor claro, es atractivo y ofrece un rendimiento coherente.|[Procedimientos recomendados para desarrollar complementos de Office](../../docs/overview/add-in-development-best-practices.md)|
|Optimice el complemento para la entrada táctil.|Haga que la interfaz de usuario responda correctamente a la entrada táctil, además de al mouse y al teclado.|
  [Aplicar principios de diseño de experiencia del usuario](https://msdn.microsoft.com/EN-US/library/mt590883.aspx#Anchor_3)|
|Haga que el complemento sea gratuito.|Office en iPad es un canal a través del cual puede llegar a más usuarios y promover sus servicios. Estos nuevos usuarios podrían convertirse en sus clientes.|
  [Directiva de validación 10.8](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|Haga que su complemento esté libre de transacciones comerciales.|El complemento no ha de tener compras desde la aplicación, ofertas de prueba, una interfaz de usuario que intente dirigir a pagos o vínculos en cualquier tienda en línea donde los usuarios pueden comprar o adquirir otro contenido, aplicaciones o complementos. En las páginas de la directiva de privacidad y de las condiciones de uso tampoco se pueden incluir interfaces de usuario comerciales o vínculos a tiendas.|
  [Directiva de validación 3.4](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|Vuelva a enviar el complemento a la Tienda Office.|En el Panel de vendedores, seleccione la casilla **Establecer este complemento como disponible en el catálogo de complementos de Office para iPad** y proporcione su identificador de desarrollador de Apple en el cuadro ID de Apple. Revise el [Contrato del proveedor de la aplicación de la Tienda Office](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.md) para asegurarse de que entiende el contrato.|
  [Enviar complementos de Office y SharePoint, y aplicaciones web de Office 365 a la Tienda Office](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)|
El complemento puede permanecer tal cual para las aplicaciones de Office que se ejecutan en otras plataformas. También puede proporcionar una interfaz de usuario diferente en función del explorador o del dispositivo en el que se ejecuta el complemento. Para detectar si el complemento se ejecuta en un iPad, puede usar las siguientes API: 

- var isTouchEnabled = [Office.context.touchEnabled](../../reference/shared/office.context.touchenabled.md)
    
- var allowCommerce = [Office.context.commerceAllowed](../../reference/shared/office.context.commerceallowed.md)
    

## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>Procedimientos recomendados para desarrollar complementos de Office para iOS y Mac

Aplique los siguientes procedimientos recomendados para desarrollar complementos que se ejecutan en iOS:


-  **Use Visual Studio para desarrollar el complemento.**
    
    Si desarrolla el complemento con Visual Studio, puede [establecer puntos de interrupción y depurar el código](../get-started/create-and-debug-office-add-ins-in-visual-studio.md#Test) en una aplicación host de Office con Windows, antes de cargar en paralelo el complemento en el iPad o Mac. Como un complemento que se ejecuta en Office para iOS u Office para Mac admite las mismas API que un complemento se ejecuta en Office para Windows, el código del complemento debe ejecutarse de la misma manera en ambas plataformas.
    
-  **Especifique los requisitos de la API en el manifiesto del complemento o con comprobaciones en tiempo de ejecución.**
    
    Al especificar los requisitos de API en el manifiesto del complemento, Office determinará si la aplicación host es compatible con los miembros de la API. Si los miembros de la API están disponibles en el host, el complemento estará disponible en la aplicación host. Como alternativa, puede realizar una comprobación en tiempo de ejecución para determinar si un método está disponible en el host antes de usarlo en el complemento. Las comprobaciones en tiempo de ejecución garantizan que el complemento siempre esté disponible y proporcionan funcionalidad adicional si los métodos están disponibles. Para obtener más información, consulte [Especificar hosts de Office y requisitos de API](../../docs/overview/specify-office-hosts-and-api-requirements.md).
    
Consulte los procedimientos recomendados generales para el desarrollo de complementos en [Procedimientos recomendados para desarrollar complementos de Office](../../docs/overview/add-in-development-best-practices.md).


## <a name="additional-resources"></a>Recursos adicionales
<a name="bk_addresources"> </a>


- [Transferir localmente un complemento de Office a iPad y Mac](../../docs/testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [Depurar complementos de Office en dispositivos iPad y Mac](../../docs/testing/debug-office-add-ins-on-ipad-and-mac.md)
    
