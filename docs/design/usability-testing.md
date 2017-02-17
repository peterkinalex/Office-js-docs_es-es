# <a name="usability-testing-for-office-add-ins"></a>Pruebas de facilidad de uso para complementos de Office

Un diseño de complemento excelente tiene en cuenta los comportamientos del usuario. Como sus propios prejuicios influencian sus decisiones de diseño, es importante probar diseños con usuarios reales para asegurarse de que sus complementos funcionan bien para sus clientes. 

Puede ejecutar pruebas de facilidad de uso de diferentes maneras. Para muchos desarrolladores de complementos, los estudios de facilidad de uso remotos y sin moderar son la manera más rentable en tiempo y costo. Algunos servicios de pruebas conocidos facilitan este proceso; a continuación se muestran algunos ejemplos: 

 - [UserTesting.com](https://www.UserTesting.com)
 - [Optimalworkshop.com](https://www.Optimalworkshop.com)
 - [Userzoom.com](https://www.Userzoom.com)

Estos servicios de pruebas le ayudan a simplificar la creación de planes de prueba y a eliminar la necesidad de buscar participantes o moderar las pruebas. 

Solo necesita cinco participantes para revelar la mayoría de problemas de uso de su diseño. Incorpore pequeñas pruebas regularmente en su ciclo de desarrollo para asegurarse de que su producto se centra en el usuario.

> **Nota:** Recomendamos que pruebe la facilidad de uso de su complemento en varias plataformas. Para [publicar su complemento en la Tienda Office](https://msdn.microsoft.com/en-us/library/office/jj220037.aspx), este debe funcionar en todas las [plataformas que admiten los métodos que define](https://dev.office.com/add-in-availability).

## <a name="1----sign-up-for-a-testing-service"></a>1.    Registrarse en un servicio de pruebas

Para obtener más información, vea [Selecting an Online Tool for Unmoderated Remote User Testing (Seleccionar una herramienta en línea para realizar pruebas de usuario remoto sin moderar)](https://www.nngroup.com/articles/unmoderated-user-testing-tools/).

## <a name="2-develop-your-research-questions"></a>2. Desarrollar las preguntas de investigación
 
Las preguntas de investigación definen los objetivos de su investigación y guían el plan de pruebas. Sus preguntas le ayudarán a identificar a los participantes que va a contratar y las tareas que realizarán. Haga que sus preguntas de investigación sean lo más específicas que pueda. También puede buscar la respuesta de preguntas más amplias.
 
A continuación se muestran algunos ejemplos de preguntas de investigación:
  
 **Específicas**  

 - ¿Los usuarios ven el vínculo de "prueba gratuita" en la página de aterrizaje?
 - Cuando los usuarios insertan contenido desde el complemento en su documento, ¿entienden dónde se ha insertado en el documento?

**Amplias**  

 - ¿Cuáles son los puntos de dificultad principales para el usuario en nuestro complemento?
 - ¿Los usuarios entienden el significado de los iconos de nuestra barra de comandos antes de que hagan clic en ellos?
 - ¿Los usuarios pueden encontrar fácilmente el menú de configuración?

Es importante obtener datos de todo el recorrido del usuario; desde la detección del complemento hasta la instalación y uso de este. Tenga en cuenta las preguntas de investigación en que se tratan los siguientes aspectos de la experiencia de usuario del complemento:
 
 - Encontrar el complemento en la tienda
 - Decidir la instalación del complemento
 - Experiencia de la primera ejecución
 - Comandos de la cinta de opciones
 - Interfaz de usuario del complemento
 - Cómo interactúa el complemento con el espacio de documento de la aplicación de Office
 - Cuánto control tiene el usuario sobre cualquier flujo de inserción de contenido

Para obtener más información, vea [Writing Effective Questions (Escribir preguntas eficaces)](http://help.usertesting.com/customer/en/portal/articles/2077663-writing-effective-questions).
 
## <a name="3-identify-participants-to-target"></a>3. Identificar participantes de destino
 
Los servicios de pruebas remotos pueden proporcionarle control sobre muchas características de los participantes de la prueba. Piense cuidadosamente en los tipos de usuarios a los que quiere dirigirse. En las primeras fases de recopilación de datos, puede que sea mejor contratar una amplia variedad de participantes para identificar más problemas de uso obvios. Después, puede decidir dirigirse a grupos como usuarios avanzados de Office, de profesiones concretas o especificar un intervalo de edad.
 
## <a name="4-create-the-participant-screener"></a>4. Crear el discriminador de participantes
 
El discriminador es el conjunto de preguntas y requisitos que presentará a los participantes potenciales de la prueba para seleccionarlos para esta. Tenga en cuenta que los participantes de servicios como UserTesting.com tienen un interés financiero en participar en la prueba. Es una buena idea incluir preguntas con trampa en su discriminador si quiere excluir a determinados usuarios de la prueba. 
 
Por ejemplo, si quiere encontrar participantes que estén familiarizados con GitHub, para filtrar usuarios que puedan tergiversar su identidad, incluya errores en la lista de posibles respuestas.

**¿Con cuáles de los siguientes repositorios de código fuente está familiarizado?**  
 a.    SourceShelf [*Rechazar*]  
 b.    CodeContainer [*Rechazar*]  
 c.    GitHub [*Debe seleccionar*]  
 d.    BitBucket [*Puede seleccionar*]  
 e.    CloudForge [*Puede seleccionar*]  


Si está planeando probar una compilación en directo de su complemento, las siguientes preguntas pueden seleccionar usuarios que puedan hacer esto. 

   **Esta prueba necesita que tenga Microsoft PowerPoint 2016. ¿Tiene PowerPoint 2016?**  
   a.    Sí [*Debe seleccionar*]  
   b.    No [*Rechazar*]  
   c.    No lo sé [*Rechazar*]  

   **Esta prueba necesita que instale un complemento gratuito para PowerPoint 2016 y que cree una cuenta gratuita para usarlo. ¿Está dispuesto a instalar un complemento y crear una cuenta gratuita?**  
    a.    Sí [*Debe seleccionar*]  
    b.    No [*Rechazar*]  

Para obtener más información, vea [Screener Questions Best Practices (Procedimientos recomendados de preguntas del discriminador)](http://help.usertesting.com/customer/en/portal/articles/2077835-screener-question-best-practices).
 
## <a name="5-create-tasks-and-questions-for-participants"></a>5. Crear tareas y preguntas para los participantes
 
Intente priorizar lo que quiere que se pruebe, de manera que pueda limitar el número de tareas y preguntas para el participante. Algunos servicios pagan a los participantes solo por una cantidad establecida de tiempo, por lo que tiene que estar seguro de que este no se va a superar.

Intente observar los comportamientos de los participantes en lugar de preguntar sobre estos, siempre que sea posible. Si necesita preguntar sobre los comportamientos, pregunte sobre lo que los participantes han hecho en el pasado, en lugar de lo que harían en una situación. Esto tiende a proporcionar resultados más confiables.
 
El principal desafío en las pruebas sin moderar es garantizar que los participantes entiendan las tareas y los escenarios. Las instrucciones deben ser *claras y concisas*. Inevitablemente, si hay alguna posibilidad de confusión, alguien se confundirá. 

No presuponga que el usuario estará en la pantalla en la que debería estar en algún punto determinado de la prueba. Considere la posibilidad de indicarles en que pantalla necesitan estar al inicio de la próxima tarea. 

Para obtener más información, vea [Writing Great Tasks (Escribir tareas adecuadas)](http://help.usertesting.com/customer/en/portal/articles/2077824-writing-great-tasks).

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a>6. Crear un prototipo que coincida con las tareas y preguntas
 
Puede probar su complemento en directo o puede probar un prototipo. Tenga en cuenta que si quiere probar el complemento en directo, necesita seleccionar participantes que tengan Office 2016, que estén dispuestos a instalar el complemento y a registrarse para obtener una cuenta (a no ser que tenga credenciales de inicio de sesión para proporcionarles). Después, necesitará asegurarse de que hayan instalado correctamente el complemento. 

De media, se tarda unos 5 minutos en guiar a los usuarios sobre la manera de instalar un complemento. A continuación se muestra un ejemplo de pasos de instalación claros y concisos. Adapte los pasos en función de los aspectos específicos de la prueba.

**Instale el complemento (inserte aquí el nombre del complemento) para PowerPoint 2016, con las siguientes instrucciones:** 

1. Abra Microsoft PowerPoint 2016.
2. Seleccione **Presentación en blanco**.
3. Vaya a **Insertar > Mis complementos**.
5. En la ventana emergente, seleccione **Tienda**.
6. Escriba (nombre del complemento) en el cuadro de búsqueda.
7. Elija (nombre del complemento).
8. Tómese un momento para observar la página Tienda para familiarizarse con el complemento.
9. Seleccione **Agregar** para instalar el complemento.

Puede probar un prototipo en cualquier nivel de interacción y fidelidad visual. Para obtener una interactividad y una vinculación más compleja, considere la posibilidad de usar una herramienta de prototipos como [InVision](https://www.invisionapp.com). Si solo quiere probar pantallas estáticas, puede hospedar imágenes en línea y enviar a los participantes la dirección URL correspondiente, o proporcionarles un vínculo a una presentación de PowerPoint en línea. 

## <a name="7-run-a-pilot-test"></a>7. Ejecutar una prueba piloto

Puede resultar difícil obtener el prototipo y la lista de preguntas y tareas correcta. Los usuarios pueden confundirse con las tareas o pueden perderse en el prototipo. Debe ejecutar una prueba piloto con entre 1 y 3 usuarios para trabajar los problemas inevitables con el formato de la prueba. Esto ayudará a garantizar que sus preguntas son claras, que el prototipo está configurado correctamente y que está obteniendo el tipo de datos que está buscando.

## <a name="8-run-the-test"></a>8. Ejecutar la prueba

Después de que ponga en marcha la prueba, obtendrá notificaciones por correo electrónico cuando los participantes la completen. A no ser que se haya dirigido a un grupo determinado de participantes, normalmente las pruebas se completan en unas horas.

## <a name="9-analyze-results"></a>9. Analizar los resultados

Esta es la parte en la que intenta comprender los datos que ha recopilado. Mientras ve los vídeos de prueba, grabe notas sobre los problemas y los aciertos que tiene el usuario. Evite intentar interpretar el significado de los datos hasta que haya visto todos los resultados. 

Un único participante que tenga un problema de uso no es suficiente para justificar que se realice un cambio en el diseño. Dos o más participantes que tengan el mismo problema sugiere que otros usuarios de la población general también lo tendrán.

En general, tenga cuidado sobre cómo usa los datos para obtener conclusiones. No caiga en la trampa de intentar que los datos se adapten a un entorno determinado; sea honesto sobre lo que los datos realmente demuestran, desmienten o si simplemente no pueden proporcionar ninguna información al respecto. Muestre una mente abierta; el comportamiento del usuario a menudo desafía las expectativas del diseñador.
 

## <a name="additional-resources"></a>Recursos adicionales
 
 - [How to Conduct Usability Testing (Cómo realizar las pruebas de facilidad de uso)](http://whatpixel.com/howto-conduct-usability-testing/)  
 - [Best Practices (Procedimientos recomendados)](http://help.usertesting.com/customer/en/portal/articles/1680726-best-practices)  
 - [Minimizing Bias (Minimizar las diferencias)](http://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
