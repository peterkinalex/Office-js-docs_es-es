-
#<a name="use-office-ui-fabric-in-office-add-ins"></a>Usar Office UI Fabric en los complementos de Office

Si va a compilar un complemento de Office, le recomendamos usar el [Office UI Fabric](https://dev.office.com/fabric) para crear la experiencia del usuario. 

Office UI Fabric es un marco front-end para crear experiencias del usuario de Office y Office 365. Fabric proporciona componentes centrados en elementos visuales que puede extender, volver a diseñar y usar en su complemento de Office. Como Fabric usa el lenguaje de diseño de Office, los componentes de experiencia de usuario de Fabric parecen una extensión natural de Office.

Fabric consta de varios proyectos:

- **Fabric JS (recomendado)**: implementa los componentes de experiencia de usuario usando solo JavaScript. Recomendamos usar esta versión de Fabric si no se desea obtener una dependencia en el marco React.  
- **Fabric React**: implementa los componentes de experiencia de usuario mediante el marco React.
- **Fabric Core**: contiene los elementos principales del lenguaje de diseño, como los iconos, los colores, el tipo y la cuadrícula. Tanto Fabric JS como Fabric React usan Fabric Core. 

Los siguientes pasos le guiarán por los conceptos básicos para el uso de Fabric JS.  

##<a name="1-add-the-fabric-cdn-references"></a>1. Agregar las referencias a Fabric en la red CDN
Para hacer referencia a Fabric desde la red CDN, agregue el siguiente código HTML a la página.

    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.min.css">
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.components.min.css">
    <script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/js/fabric.min.js"></script>

Y eso es todo. Ya puede empezar a usar Fabric en su complemento. 

##<a name="2-use-fabric-icons-and-fonts"></a>2. Usar fuentes e iconos del Tejido
Usar iconos es sencillo. Lo único que debe hacer es usar un elemento "i" y hacer referencia a las clases correspondientes. Puede controlar el tamaño del icono cambiando el tamaño de fuente. Por ejemplo, en el código siguiente, se muestra cómo crear un icono de tabla muy grande que use el color themePrimary (#0078d7). 
   
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>

Para buscar más iconos disponibles en Office UI Fabric, use la característica de búsqueda de la página [Iconos](https://dev.office.com/fabric#/styles/icons). Cuando encuentre un icono para usar en su complemento, asegúrese de agregar `ms-Icon--` al nombre como prefijo. 

Para obtener información acerca de los tamaños de fuente y los colores que están disponibles en Office UI Fabric, consulte [Tipografía](https://dev.office.com/fabric#/styles/typography) y [Colores](https://dev.office.com/fabric#/styles/colors).

##<a name="3-use-fabric-js-ux-components"></a>3. Usar componentes de experiencia de usuario de Fabric JS

Fabric proporciona varios componentes de experiencia de usuario, como botones o casillas, que puede usar en el complemento. La siguiente es una lista de los componentes de experiencia de usuario de Fabric JS que recomendamos usar en un complemento. Para usar alguno de los componentes de Fabric en el complemento, siga el vínculo a la documentación de Fabric y, a continuación, siga las instrucciones de la sección **Uso de este componente**.

> **Nota:** Agregaremos componentes adicionales con el tiempo. 

- [Ruta de navegación](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Breadcrumb.md)
- [Botón](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Button.md)
- [Casilla](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/CheckBox.md)
- [ChoiceFieldGroup](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/ChoiceFieldGroup.md)
- [Selector de fecha](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/DatePicker.md) (para obtener un ejemplo que muestre cómo implementar el selector de fecha en un complemento, consulte el ejemplo de código [Seguimiento de ventas de Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)).
- [Lista desplegable](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Dropdown.md)
- [Etiqueta](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Label.md)
- [Vínculo](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Link.md)
- [Lista](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/List.md) (considere la posibilidad de cambiar los estilos predeterminados del componente por defecto en las CSS).
- [MessageBanner](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/MessageBanner.md)
- [MessageBar](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/MessageBar.md)
- [Superposición](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Overlay.md)
- [Panel](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Panel.md)
- [Barra dinámica](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Pivot.md)
- [ProgressIndicator](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/ProgressIndicator.md)
- [Cuadro de búsqueda](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/SearchBox.md)
- [Control de giro](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Spinner.md)
- [Tabla](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Table.md)
- [TextField](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/TextField.md)
- [Botón de alternancia](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Toggle.md)
   
## <a name="updating-your-add-in-to-use-fabric-js"></a>Actualizar el complemento para usar Fabric JS
Si ha estado usando una versión anterior de Office UI Fabric y le gustaría actualizar a Fabric JS, asegúrese de obtener información acerca de los nuevos componentes, de incorporarlos y de probarlos en su complemento. Tenga en cuenta las siguientes cuestiones a modo de ayuda para planear sus actualizaciones:

- La inicialización de componentes es más sencilla con Fabric JS. En las versiones anteriores de Fabric, deberá incluir el archivo JavaScript del componente de Fabric en su proyecto de complemento, incluida una referencia de `<Script>` a ese archivo y, a continuación, inicializar el componente. En Fabric JS, ya no es necesario que incluya el archivo JavaScript del componente de Fabric ni la referencia de `<Script>` asociada. Todo lo que debe hacer es inicializar el componente de Fabric.   
- Ahora, hay varios componentes que proporcionan funciones que controlan el comportamiento del componente de experiencia de usuario. Por ejemplo, el control de casilla tiene una función de `toggle` que permite seleccionar y desactivar. 
- Se han actualizado estilos y nombres de clase de algunos iconos.
- El cambio más notable es el uso del elemento `<label>` en muchos componentes. El elemento `<label>` controla el estilo del componente. Puede que necesite actualizar su código de experiencia de usuario para poder usar el elemento `<label>`. Por ejemplo, cambiar el valor del atributo seleccionado del elemento `<input>` en una casilla de Fabric JS no tendrá ningún efecto en la casilla. En lugar de eso, use las funciones `check`, `unCheck` o `toggle`.   

##<a name="next-steps"></a>Pasos siguientes
Si busca un ejemplo completo en el que se muestre cómo usar Fabric JS, tenemos justo lo que necesita. Consulte el siguiente recurso:

- [Seguimiento de ventas de Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 

##<a name="related-resources"></a>Recursos relacionados
Si busca ejemplos de código o documentación acerca de alguna versión anterior de Fabric, consulte lo siguiente:

- [Modelos de diseño de la experiencia del usuario (usa Fabric 2.6.1)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 
- [Ejemplo de la interfaz de usuario de Fabric en un complemento de Office (usa Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) 
- [Usar Fabric 2.6.1 en un complemento de Office](https://dev.office.com/docs/add-ins/design/ui-elements/using-office-ui-fabric)
 

