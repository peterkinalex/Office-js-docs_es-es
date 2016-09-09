
# Patrones de interacción de los complementos de Office


Los Complementos de Office pueden mejorar las experiencias de creación y productividad, así como conectar el contenido de aplicaciones de host de Office a flujos de trabajo mayores basados en web. Hay una serie de escenarios comunes que se aplican a los complementos de contenido, panel de tareas y Outlook que se pueden desarrollar. En este artículo se describen algunos de los escenarios más habituales y se ofrecen patrones de interacción recomendados para la experiencia de usuario del complemento. Puede desglosar, combinar o mezclar y emparejar estos patrones de interacción en función de sus escenarios exclusivos.

 **Escenarios de complementos comunes**

| Tipo de complemento | Escenarios comunes |
| ------ | ------ |
|  Contenido  |  Visualizar datos <br> Widgets y herramientas  |
|  Panel de tareas  |  Transformación y procesamiento de datos <br> Creación de forma eficaz y eficiente <br> Búsqueda de contenido e inserción de datos <br> Publicar o cargar contenido a un servicio web  |
|  Outlook  |  Hacer de puente entre el contenido de correo y una aplicación externa <br> Dar más información sobre el contenido en una cita o mensaje de correo <br> Proporcionar información que le ayude a ser más productivo  |

## Visualizar datos con un complemento de contenido


Este ejemplo muestra un complemento de contenido para Excel que genera un gráfico a partir de los datos de una hoja de cálculo.

En este patrón de interacción, el complemento no se vuelve activo hasta que se selecciona y se enlazan los datos para generar el gráfico. Es importante comunicar el propósito del complemento y los pasos para activarlo en la vista inicial del complemento. 

**Complemento de contenido para Excel que genera un gráfico a partir de los datos de una hoja de cálculo.**
<br>
![Aplicación de contenido para Excel que genera un gráfico a partir de los datos de una hoja de cálculo](../../../images/off15appUXFig01.png)
<br>
<ul><li><p>Muestra instrucciones junto con un botón deshabilitado (A) para reforzar que es necesario realizar una acción antes de elegir el botón.</p></li><li><p>Después de seleccionar un rango de celdas, el botón <span class="ui">Crear gráfico</span> se activa (B - C).</p></li><li><p>La visualización se propagará en el contenedor y reemplazará a la vista anterior (D).</p></li><li><p>Muestre cualquier UI adicional en el borde inferior del complemento, junto con un botón de configuración (un engranaje) que le llevará a una vista donde podrá restablecer o administrar el complemento.</p></li></ul>Recomendado para:
<ul><li><p>Complementos que requieren que seleccione datos antes de la activación.</p></li></ul>

## Transformar contenido con un complemento de panel de tareas


Este ejemplo muestra un complemento de panel de tareas que traduce el texto del documento a otro idioma.

En este patrón de interacción, primero hay que seleccionar el texto del documento que se quiere traducir.

**Complemento de panel de tareas que traduce el texto del documento a otro idioma.**
<br>
![Aplicación de panel de tareas que traduce el texto del documento a otro idioma](../../../images/off15appUXFig02.png)
<br>
<ul><li><p>Comunique el propósito del complemento en un título y una sugerencia en el hecho de que primero debe realizar una selección (A).</p></li><li><p>El menú de idioma y el botón <span class="ui">Traducir</span> están deshabilitados, lo que refuerza el hecho de que debe realizar una acción para poder continuar. Después de seleccionar contenido en el documento, estos dos elementos se activarán (D).</p></li><li><p>Después de elegir <span class="ui">Traducir</span>, la interfaz de usuario se despliega para mostrar el contenido traducido junto con un botón que permite insertarlo de nuevo en el documento (E).</p></li><li><p>Puede proporcionar un botón <span class="ui">Borrar</span> o <span class="ui">Restablecer</span> que vuelva a la vista inicial.</p></li></ul>Recomendado para:
<ul><li><p>Complementos que requieren que seleccione datos antes de la activación.</p></li><li><p>Interfaz de usuario que se despliega o que se muestra conforme va progresando durante el escenario.</p></li></ul>

## Procesar datos con un complemento de panel de tareas


Este ejemplo muestra un complemento de panel de tareas que comprueba datos en Excel.

En este patrón de interacción, hay que seleccionar un rango de celdas en la hoja de cálculo para comenzar.

**Complemento de panel de tareas que comprueba datos en Excel**
<br>
![Aplicación de panel de tareas que comprueba datos en Excel](../../../images/off15appUXFig03.png)
<br>
<ul><li><p>El propósito del complemento se describe en el título. Las instrucciones le ayudan a empezar.</p></li><li><p>El botón <span class="ui">Enviar datos seleccionados</span> está deshabilitado, lo que refuerza el hecho de que debe llevar a cabo una acción para poder continuar (A).</p></li><li><p>Después de seleccionar un rango de celdas de la hoja de cálculo (B), el botón <span class="ui">Enviar datos seleccionados</span> se activa.</p></li><li><p>Después de elegir este botón, la interfaz de usuario se reemplaza por el rango de celdas seleccionado, una barra de progreso y un botón <span class="ui">Cancelar</span>.</p></li><li><p>La barra de progreso indica el estado del proceso y el botón <span class="ui">Cancelar</span> le permite interrumpirlo (D).</p></li><li><p>Una vez finalizado el proceso, los resultados se muestran automáticamente (E). Al seleccionar un elemento en la lista, se activa la celda correspondiente en la hoja de cálculo.</p></li></ul>Recomendado para:
<ul><li><p>Procesos cuya duración es indeterminada.</p></li></ul>

## Analizar contenido con un complemento de panel de tareas


Este ejemplo muestra un complemento de panel de tareas que muestra definiciones de palabras conforme las escribe.

En este patrón de interacción, hay que seleccionar primero el texto en el documento para ver los resultados.

**Complemento de panel de tareas que muestra definiciones de palabras conforme las escribe**
<br>
![Aplicación de panel de tareas que muestra definiciones de palabras conforme el usuario escribe](../../../images/off15appUXFig04.png)
<br>
<ul><li><p>Un titular explica la finalidad del complemento y cómo empezar (A).</p></li><li><p>La búsqueda automática se activa de forma predeterminada con la opción para deshabilitarla (B).</p></li><li><p>Después de hacer una selección, el complemento muestra el contenido correspondiente (D).</p></li><li><p>Proporcione un vínculo para mostrar más información (E).</p></li></ul>Recomendado para:
<ul><li><p>Complementos que devuelven datos automáticamente conforme crea contenido.</p></li><li><p>Complementos que requieren que seleccione contenido antes de la activación.</p></li></ul>

## Buscar contenido con un complemento de panel de tareas


Este ejemplo muestra un complemento de panel de tareas para buscar contenido.

En este patrón de interacción, se escribe una cadena en el cuadro de búsqueda o se selecciona una opción en una lista de contenido destacado para comenzar.

**Complemento de panel de tareas para la búsqueda de contenido**
<br>
![Aplicación de panel de tareas para la búsqueda de contenido](../../../images/off15appUXFig05.png)
<br>
<ul><li><p>La vista inicial contiene un cuadro de <span class="ui">búsqueda</span> (A) y una lista de contenido presentado (B).</p></li><li><p>Cuando escribe una cadena en el cuadro de búsqueda, el icono de búsqueda se reemplaza por un icono de cierre (C).</p></li><li><p>Si elige el icono de cierre, vuelve a la vista inicial.</p></li></ul>Recomendado para:
<ul><li><p>Complementos que devuelven datos automáticamente conforme crea contenido.</p></li><li><p>Complementos que requieren que seleccione contenido antes de la activación.</p></li></ul>

## Insertar elementos multimedia con un complemento de panel de tareas


En este patrón de interacción, se puede seleccionar una imagen en los resultados de búsqueda para insertarla en su documento.

**Complemento de panel de tareas para insertar una imagen**
<br>
![Aplicación de panel de tareas para insertar una imagen](../../../images/off15appUXFig06.png)
<br>
<ul><li><p>Ha filtrado una lista de resultados de búsqueda (A) y el contenido seleccionado para insertar (B).</p></li><li><p>Se muestra una vista detallada del contenido seleccionado (C) junto con un botón que le permite volver a la lista.</p></li><li><p>En el pie de página hay un botón para <span class="ui">insertar foto</span> (D). Después de elegir este botón, la imagen se inserta en el documento.</p></li><li><p>Junto con el contenido insertado, se incluye una breve descripción del lugar del que procede la imagen (E). </p></li><li><p>La interfaz de usuario del complemento confirma visualmente que la acción se ha llevado a cabo correctamente.</p></li></ul>Recomendado para:
<ul><li><p>Complementos de inserción de contenido.</p></li></ul>

## Insertar texto seleccionado con un complemento de panel de tareas


En este patrón de interacción, se selecciona texto en los resultados de búsqueda para insertarlo en el documento.

**Complemento de panel de tareas para insertar texto**
<br>
![Aplicación de panel de tareas para insertar texto](../../../images/off15appUXFig07.png)
<br>
<ul><li><p>Ya ha encontrado un fragmento de contenido (A).</p></li><li><p>En el pie de página hay un botón deshabilitado para <span class="ui">insertar selección</span> (B).</p></li><li><p>Cuando seleccione una cadena de texto (C), el botón para <span class="ui">insertar selección</span> se activa.</p></li><li><p>Después de elegir este botón, el complemento inserta el texto seleccionado en el documento junto con una referencia al origen del contenido (E).</p></li></ul>Recomendado para:
<ul><li><p>Complementos para llevar a cabo búsquedas e insertar contenido.</p></li></ul>

## Publicar en un servicio web con un complemento de panel de tareas


Este ejemplo muestra un complemento de panel de tareas para publicar un documento como una publicación de blog.

En este patrón de interacción, se ha completado la escritura de contenido en un documento y se quiere publicar en un blog.

**Complemento de panel de tareas para la publicación de un documento como publicación de blog.**
<br>
![Aplicación de panel de tareas para la publicación de un documento como entrada de blog](../../../images/off15appUXFig08.png)
<br>
<ul><li><p>En primer lugar, se muestra un formulario de inicio de sesión para escribir sus credenciales (A).</p></li><li><p>Se proporcionan vínculos para crear una cuenta y tratar los problemas de inicio de sesión más típicos (B). Si elige estos vínculos, se abren estas páginas en un explorador.</p></li><li><p>Después de haber iniciado sesión, el complemento muestra un formulario para crear una nueva entrada de blog (C).</p></li><li><p>Hacia la parte superior del complemento se muestra el nombre de la cuenta en la que ha iniciado sesión (y en la que realizará la publicación). El complemento proporciona un vínculo para obtener una vista previa de la publicación (D). Si elige este vínculo, se muestra la vista previa en un explorador.</p></li><li><p>Después de elegir la opción para <span class="ui">crear publicación</span>, el complemento muestra una vista donde se confirma que el contenido del documento se ha publicado (E).</p></li><li><p>El complemento incluye un vínculo para ver la publicación en un explorador (F), además de un botón para crear otra publicación (G).</p></li></ul>Recomendado para:
<ul><li><p>Complementos que muestran, publican o comparten contenido en redes sociales, sitios de blogs y servicios web.</p></li><li><p>Complementos que requieren que inicie sesión en un servicio.</p></li></ul>

## Obtener más información sobre personas con un complemento de Outlook


 **Ejemplo 1**

**Página de resultados y detalles**
<br>
![Página de resultados y detalles](../../../images/off15appUXFig09.jpg)
<br>
Recomendado para:
<ul><li><p>Ver la amplitud del contenido en caso de que disponga de grandes conjuntos de datos útiles para mostrar.</p></li><li><p>Páginas de detalles que usan la capacidad completa del contenedor del complemento.</p></li><li><p>Modelos de navegación que usan un flujo "hacia adelante y hacia atrás".</p></li></ul>
 **Ejemplo 2**

**Página de detalles con navegación persistente**
<br>
![Página de detalles con navegación persistente](../../../images/off15appUXFig10.jpg)
<br>
Recomendado para:
<ul><li><p>Visualización de forma predeterminada del primer resultado de un conjunto de datos.</p></li><li><p>Estructuras de navegación que funcionan de forma similar a las pestañas (navegación lineal de un único nivel).</p></li><li><p>Reducción de acciones de selección al mantener la navegación disponible en todo momento.</p></li><li><p>Espacio para mostrar la navegación en todo momento.</p></li></ul>

## Obtener más información sobre contenido con un complemento de Outlook


 **Ejemplo 1**

**Página de resultados y detalles**
<br>
![Página de resultados y detalles](../../../images/off15appUXFig11.jpg)
<br>
Recomendado para:
<ul><li><p>Ver la amplitud del contenido en caso de que disponga de grandes conjuntos de datos útiles para mostrar.</p></li><li><p>Pedirle que elija una opción o selección antes de mostrar más detalles.</p></li><li><p>Páginas de detalles que usan la capacidad completa del contenedor del complemento.</p></li><li><p>Modelos de navegación que usan un flujo "hacia adelante y hacia atrás".</p></li></ul>
 **Ejemplo 2**

**Página de detalles con contenido secundario**
<br>
![Página de detalles con contenido secundario](../../../images/off15appUXFig12.jpg)
<br>
Recomendado para:
<ul><li><p>Casos en los que desee incluir parte de un contenido.</p></li><li><p>Mostar el contenido sin la interacción del usuario.</p></li><li><p>Navegación persistente (puede agregarse a este modelo para obtener una mezcla de sencillez y facilidad de navegación).</p></li></ul>

## Conectar con un servicio en línea y presentar datos


Estos ejemplos muestran patrones de interacción para obtener datos y contenido de un servicio en línea. Se pueden usar en los tres tipos de complementos: complementos de contenido, complementos de panel de tareas y complementos de Outlook.

 **Ejemplo 1**

**Carrusel**
<br>
![Carrusel](../../../images/off15appUXFig13.jpg)
<br>
Recomendado para:
<ul><li><p>Pequeñas cantidades de datos que pueden mostrarse de una en una o en grupos.</p></li><li><p>Visualización de contenido en formato de galería como, por ejemplo, en presentaciones de diapositivas o galerías de imágenes.</p></li><li><p>Cuando el modelo de navegación hacia adelante/atrás funciona correctamente.</p></li></ul>
 **Ejemplo 2**

**Asistente**
<br>
![Asistente](../../../images/off15appUXFig14.jpg)
<br>
Recomendado para:
<ul><li><p>Contenido que debe mostrarse con un orden determinado.</p></li><li><p>Grandes cantidades de contenido que deben mostrarse en fragmentos de información de menor tamaño.</p></li><li><p>Experiencias en las que se muestra el contenido a los usuarios en forma de libro.</p></li><li><p>Cuando es necesario llevar a cabo una serie de pasos o acciones para completar una tarea.</p></li></ul>
 **Ejemplo 3**

**Formulario, resultados y detalles**
<br>
![Formulario, resultados y detalles](../../../images/off15appUXFig15.jpg)
<br>
Recomendado para:
<ul><li><p>Complementos que necesitan entrada de datos.</p></li><li><p>Punto de entrada para mostrar resultados y patrones de detalles.</p></li></ul>

## Recursos adicionales



- [Directrices de diseño para complementos de Office](../add-in-design.md)
    
