# <a name="outlook-add-in-design-guidelines"></a>Instrucciones de diseño del complemento de Outlook

Los complementos son una gran manera de que los partners extiendan las funciones de Outlook más allá de nuestro conjunto principal de características. Los complementos permiten a los usuarios tener acceso a contenido, tareas y experiencias de terceros sin salir de su bandeja de entrada. Una vez instalados, los complementos de Outlook están disponibles en cualquier plataforma y dispositivo. Las siguientes instrucciones de alto nivel le ayudarán a diseñar y compilar un complemento atractivo, que proporcione lo mejor de su aplicación directamente en Outlook; en Windows, Web, iOS, Mac y Android (próximamente).

## <a name="principles"></a>Principios

1. **Céntrese en unas pocas tareas clave; hágalo correctamente**

    Los complementos mejor diseñados son sencillos de usar, específicos y proporcionan un valor real a los usuarios. Como su complemento se ejecutará dentro de Outlook, se hace un énfasis adicional en este principio. Outlook es una aplicación de productividad; es el lugar donde los usuarios acuden para realizar las cosas.

    Será una extensión de nuestra experiencia y es importante asegurarse de que los escenarios que habilite se adapten de manera natural dentro de Outlook. Piense cuidadosamente en qué casos de uso comunes se aprovecharán más de tener enlaces a ellos desde nuestras experiencias de calendario y correo electrónico.

    Un complemento no debe intentar realizar todo lo que hace la aplicación. Debe centrarse en las acciones más apropiadas y que más se usan en el contexto del contenido de Outlook. Piense en su llamada a la acción y deje claro lo que debe hacer el usuario cuando se abra el panel de tareas.

2. **Haga que se sienta tan nativo como sea posible**

    El complemento debe diseñarse mediante patrones nativos para la plataforma en la que Outlook se está ejecutando. Para conseguirlo, asegúrese de respetar e implementar la interacción y las instrucciones visuales que establece cada plataforma. Outlook tiene sus propias instrucciones y también es importante tenerlas en cuenta. Un complemento bien diseñado será una combinación apropiada de su experiencia, de la plataforma y de Outlook.

    Significa que el complemento tendrá que ser visualmente diferente cuando se ejecute en Outlook para iOS y en Outlook para Android (cuando lancemos la compatibilidad para dicha versión). Es recomendable echar un vistazo a [Framework7](https://framework7.io/) para obtener ayuda con el estilo. Publicaremos pautas actualizadas, sobre todo para Android, conforme nos acerquemos al lanzamiento de la compatibilidad con complementos de Outlook para Android.

3. **Haga que sea agradable de usar y obtenga los detalles correctos**

    Los usuarios disfrutan usando productos que son atractivos a nivel funcional y visual. Puede ayudar a garantizar el éxito de su complemento diseñando una experiencia en la que haya considerado cuidadosamente cada detalle visual y de interacción. Los pasos necesarios para completar una tarea deben ser claros y relevantes. Idealmente, ninguna acción debe necesitar más de uno o dos clics. Intente no sacar al usuario de contexto para completar una acción. Un usuario debe poder entrar y salir fácilmente del complemento y volver a lo que estuviera realizando anteriormente. Un complemento no pretende ser un destino en el que invertir mucho tiempo; es una mejora de nuestras características principales. Si se ha realizado correctamente, el complemento nos ayudará a cumplir con el objetivo de que los usuarios sean más productivos.

4. **Proporcione una marca de manera acertada**

    Valoramos una gran personalización de marca y sabemos que es importante proporcionar a los usuarios su experiencia única. Pero creemos que la mejor manera de garantizar el éxito del complemento es crear una experiencia intuitiva que incorpore sutilmente elementos de su marca, en lugar de mostrar elementos de marca constantes y molestos que solo distraen al usuario de moverse por el sistema de una manera libre. Una buena manera de incorporar su marca de manera significativa es mediante el uso de su identidad cromática, iconos y voz; si se presupone que estos no entran en conflicto con los requisitos de accesibilidad ni con los patrones de plataforma preferidos. Intente centrarse en el contenido y la finalización de tareas, no en la atención de la marca.

## <a name="design-patterns"></a>Patrones de diseño

> **Nota:** Aunque los principios anteriores se aplican a todos los puntos de conexión y plataformas, los siguientes patrones y ejemplos son específicos de complementos móviles en la plataforma de iOS.

Para ayudarle a crear un complemento bien diseñado, tenemos [plantillas](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/tree/master/Helpful%20Templates/Outlook%20Mobile) que contienen patrones móviles de iOS que funcionan dentro del entorno de Outlook Mobile. Aprovechar estos patrones específicos ayudará a garantizar que su complemento se sienta nativo en la plataforma de iOS y en Outlook Mobile. Estos patrones también se detallan a continuación. Aunque no es exhaustivo, este es el comienzo de una biblioteca que seguiremos creando hasta que revelemos partners de paradigmas adicionales que quieran incluir en sus complementos.  

### <a name="overview"></a>Información general

Un complemento típico incluye los componentes siguientes.

![Un diagrama de patrones de UX básicos para un panel de tareas en iOS](../../images/outlook-mobile-design-overview.png)

### <a name="loading"></a>Carga

Cuando un usuario pulsa en el complemento, la UX debe mostrarse tan rápido como sea posible. Si existe algún retraso, use una barra de progreso o un indicador de actividad. Una barra de progreso debe usarse cuando la cantidad de tiempo se puede determinar y un indicador de actividad debe usarse cuando la cantidad de tiempo es indeterminable.

![Ejemplos de una barra de progreso y un indicador de actividad en iOS](../../images/outlook-mobile-design-loading.png)

### <a name="sign-insign-up"></a>Iniciar sesión y registrarse

Haga que su inicio de sesión (y registro) fluya fácilmente y sea sencillo de usar.

![Ejemplos de páginas de registro e inicio de sesión en iOS](../../images/outlook-mobile-design-signin.png)

### <a name="brand-bar"></a>Barra de marca

La primera pantalla de su complemento debe incluir su elemento de personalización de marca. Diseñada para el reconocimiento, la barra de marca también ayuda a establecer el contexto para el usuario. Como la barra de navegación contiene el nombre de su empresa o marca, no es necesario repetir la barra de marca en las páginas posteriores.

![Ejemplos de barras de marca en iOS](../../images/outlook-mobile-design-branding.png)

### <a name="margins"></a>Márgenes

Los márgenes móviles deben establecerse en 15 px (8 % de la pantalla) para cada lado, para alinearse con Outlook para iOS.

![Ejemplos de márgenes en iOS](../../images/outlook-mobile-design-margins.png)

### <a name="typography"></a>Tipografía

El uso de tipografía se alinea con Outlook para iOS y se mantiene de manera simple para su legibilidad.

![Ejemplos de tipografía para iOS](../../images/outlook-mobile-design-typography.png)

### <a name="color-palette"></a>Paleta de colores

El uso de colores es sutil en Outlook para iOS.  Para alinear, le pedimos que el uso de color se localice en acciones y estados de error, de manera que solo la barra de marca use un color único.

![Paleta de colores para iOS](../../images/outlook-mobile-design-color-palette.png)

### <a name="cells"></a>Celdas

Como la barra de navegación no puede usarse para etiquetar una página, use títulos de sección para etiquetar páginas.

![Tipos de celda para iOS](../../images/outlook-mobile-design-cell-types.png)
* * *
![Tipos de celda "que se deben hacer" para iOS](../../images/outlook-mobile-design-cell-dos.png)
* * *
![Tipos de celda "que no se deben hacer" para iOS](../../images/outlook-mobile-design-cell-donts.png)
* * *
![Celdas y entradas para iOS](../../images/outlook-mobile-design-cell-input.png)

### <a name="actions"></a>Acciones

Incluso si su aplicación controla una multitud de acciones, piense en las más importantes que quiere que realice el complemento y concéntrese en ellas.

![Acciones y celdas en iOS](../../images/outlook-mobile-design-action-cells.png)
* * *
![Acciones "que se deben hacer" para iOS](../../images/outlook-mobile-design-action-dos.png)

### <a name="buttons"></a>Botones

Los botones se usan cuando existen otros elementos de UX debajo (frente a las acciones, donde la acción es el último elemento de la pantalla).

![Ejemplos de botones para iOS](../../images/outlook-mobile-design-buttons.png)

### <a name="tabs"></a>Pestañas

Las pestañas pueden ayudar en la organización de contenido.

![Ejemplos de pestañas para iOS](../../images/outlook-mobile-design-tabs.png)

### <a name="icons"></a>Iconos

Los iconos deben seguir el diseño actual de Outlook para iOS cuando sea posible. Use nuestro tamaño y color estándar.

![Ejemplos de iconos para iOS](../../images/outlook-mobile-design-icons.png)

## <a name="end-to-end-examples"></a>Ejemplos completos

Para el lanzamiento de nuestros complementos de Outlook Mobile v1, hemos trabajado estrechamente con nuestros partners que estaban creando complementos. A modo de muestra del potencial de sus complementos en Outlook Mobile, nuestro diseñador ha reunido flujos completos para cada complemento, para aprovechar nuestras instrucciones y patrones.

> **Nota importante:** Estos ejemplos pretenden destacar la manera ideal de enfocar el diseño visual y de interacción de un complemento y puede que no coincidan con los conjuntos exactos de características de las versiones enviadas de los complementos. 

### <a name="giphy"></a>GIPHY

![Diseño completo del complemento de GIPHY](../../images/outlook-mobile-design-giphy.png)

### <a name="nimble"></a>Nimble

![Diseño completo del complemento de Nimble](../../images/outlook-mobile-design-nimble.png)

### <a name="trello"></a>Trello

![Parte 1 del diseño completo del complemento de Trello](../../images/outlook-mobile-design-trello-1.png)
* * *
![Parte 2 del diseño completo del complemento de Trello](../../images/outlook-mobile-design-trello-2.png)
* * *
![Parte 3 del diseño completo del complemento de Trello](../../images/outlook-mobile-design-trello-3.png)

### <a name="dynamics-crm"></a>Dynamics CRM

![Diseño completo del complemento de Dynamics CRM](../../images/outlook-mobile-design-crm.png)
