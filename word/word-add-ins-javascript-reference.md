# Referencia de JavaScript de complementos de Word 

Busque la referencia de API para la API de JavaScript de Word para los complementos de Word.

_Se aplica a: Word 2016, Word para iPad, Word para Mac_

## En esta sección

Estos son los objetos principales de la API de JavaScript de Word.

* [Body](word-add-ins-javascript-reference/body.md): representa el cuerpo de un documento o una sección.
* [ContentControl](word-add-ins-javascript-reference/contentcontrol.md): contenedor para el contenido. Se trata de una región delimitada y posiblemente con etiquetas en un documento que sirve como contenedor para tipos de contenido específicos. Por ejemplo, ContentControl puede contener contenido como párrafos de texto con formato y otros controles de contenido. Puede acceder a un control de contenido a través de la colección de controles de contenido del documento, el cuerpo del documento, el párrafo, el intervalo o un control de contenido.
* [Document](word-add-ins-javascript-reference/document.md): objeto de nivel superior. Un objeto Document contiene una o más [secciones](word-add-ins-javascript-reference/section.md), un cuerpo que contiene el contenido del documento e información de encabezado y pie de página.
* [Font](word-add-ins-javascript-reference/font.md): proporciona formato de texto a un cuerpo, control de contenido, párrafo o intervalo.
* [Image](word-add-ins-javascript-reference/inlinepicture.md): representa una imagen incorporada anclada a un párrafo.
* [Paragraph](word-add-ins-javascript-reference/paragraph.md): representa un solo párrafo de una selección, intervalo o documento. Puede acceder a un párrafo a través de la colección de párrafos de una selección, intervalo o documento. 
* [Intervalo](word-add-ins-javascript-reference/range.md): representa un área contigua en un documento. Se obtiene un objeto Range cuando se obtiene una selección, se inserta contenido en el cuerpo, se inserta contenido en un control de contenido, se inserta contenido en un párrafo o se obtiene un resultado de búsqueda. Puede definir y manipular un intervalo sin cambiar la selección.
* [Section](word-add-ins-javascript-reference/section.md):  define los diferentes encabezados y pies de página, así como las diferentes configuraciones del diseño de página de un documento. Puede acceder a las secciones desde el objeto Document. 
* [Selection](word-add-ins-javascript-reference/document.md#getselection): el objeto Document proporciona acceso a la selección del usuario en el documento o al punto de inserción actual si no hay nada seleccionado.

## Denos su opinión

Su opinión es importante para nosotros. 

* Consulte los documentos y háganos saber todas las preguntas y las dificultades que le planteen [enviando un problema](https://github.com/OfficeDev/office-js-docs/issues) directamente en este repositorio.
* Infórmenos sobre su experiencia de programación, lo que le gustaría ver en versiones futuras, ejemplos de código, etc. Use [este sitio](http://officespdev.uservoice.com/) para enviar sus sugerencias e ideas.

## Recursos adicionales

* [Complementos de Word](word-add-ins.md)
* [Guía de programación de complementos de Word](word-add-ins-programming-guide.md)
* [Complementos de Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Introducción a los complementos de Office](http://dev.office.com/getting-started/addins)
* &lt;a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=Word"&gt;Complementos de Word en GitHub&lt;/a&gt;
* [Explorador de fragmentos de código para Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
