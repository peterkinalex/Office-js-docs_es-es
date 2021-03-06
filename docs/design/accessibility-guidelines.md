# <a name="accessibility-guidelines-for-office-add-ins"></a>Directrices de accesibilidad para complementos de Office

Al diseñar y desarrollar complementos de Office, le conviene asegurarse de que todos los usuarios y los clientes potenciales puedan usar correctamente su complemento. Aplique las directrices siguientes para asegurarse de que la solución sea accesible para todas las audiencias.

##<a name="design-for-multiple-input-methods"></a>Diseñar para varios métodos de entrada

- Asegúrese de que los usuarios pueden realizar las operaciones usando solo el teclado. Los usuarios deben ser capaces de mover todos los elementos de la página que requieren una acción mediante una combinación de las teclas TAB y de flecha.
- En un dispositivo móvil, cuando los usuarios accionan un control con el tacto, el dispositivo debe proporcionar respuestas de audio útiles.
- Proporcione etiquetas útiles para todos los controles interactivos. 

##<a name="make-your-add-in-easy-to-use"></a>Hacer que el complemento sea fácil de usar

- No confíe en un solo atributo (como el color, el tamaño, la forma, la ubicación, la orientación o el sonido) para transmitir un propósito en la interfaz de usuario.
- Evite los cambios de contexto inesperados, como centrarse en otro elemento de la interfaz de usuario sin la intervención del usuario.
- Proporcione una manera de comprobar, confirmar o revertir todas las acciones de enlace.
- Proporcione una manera de pausar o detener medios como audio y vídeo.
- No imponga un límite de tiempo para acciones del usuario.

##<a name="make-your-add-in-easy-to-see"></a>Hacer que el complemento sea fácil de ver

- Evite los cambios de color inesperados.
- Proporcione información útil y oportuna para describir elementos de la interfaz de usuario, títulos y encabezados, entradas y errores. Asegúrese de que los nombres de los controles describen adecuadamente el propósito del control.
- Siga las [directrices estándar](http://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) para el contraste de colores.

##<a name="account-for-assistive-technologies"></a>Tener en cuenta las tecnologías de ayuda

- Evite usar funciones que interfieran con las tecnologías de ayuda, incluidas las interacciones visuales, auditivas o de otro tipo.
- No proporcione texto en formato de imagen. Los lectores de pantalla no pueden leer el texto contenido en imágenes.
- Proporcione una manera de que los usuarios ajusten o silencien todos los orígenes de audio.
- Proporcione una manera de que los usuarios activen la descripción de títulos o audio con orígenes de audio.
- Proporcione alternativas al sonido como medio para alertar a los usuarios, como indicaciones visuales o vibraciones.

##<a name="accessibility-resources"></a>Recursos de accesibilidad

- [Directrices de accesibilidad para contenido web (WCAG) 2.0](http://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [Guía para aplicar las WCAG 2.0 a información que no pertenece a la Web y a tecnologías de la información (WCAG2ICT)](http://www.w3.org/TR/wcag2ict/)
- [Estándar europeo sobre los requisitos de accesibilidad de las tecnologías de la información (TIC)](http://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf)


