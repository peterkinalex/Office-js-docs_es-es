# Referencia de la API de JavaScript de complementos de OneNote

*Válido para: OneNote Online*

En los vínculos siguientes se muestran los objetos de OneNote de alto nivel disponibles en la API. Cada vínculo a una página del objeto contiene una descripción de las propiedades, relaciones y métodos disponibles en el objeto. Explore los vínculos siguientes para más información. 
    
- [Application](application.md): el objeto de nivel superior que se usa para tener acceso a todos los objetos de OneNote a los que se puede hacer referencia globalmente, el bloc de notas activo y la sección activa.

- [Notebook](notebook.md): un bloc de notas. Los blocs de notas contienen grupos de secciones y secciones.

   - [NotebookCollection](notebookcollection.md): una colección de blocs de notas.

- [SectionGroup](sectiongroup.md): un grupo de secciones. Los grupos de secciones contienen grupos de secciones y secciones.

   - [SectionGroupCollection](sectiongroupcollection.md): una colección de grupos de secciones.

- [Section](section.md): una sección. Las secciones contienen páginas.

   - [SectionCollection](sectioncollection.md): una colección de secciones.

- [Page](page.md): una página. Las páginas contienen objetos PageContent.

   - [PageCollection](pagecollection.md): una colección de páginas.

- [PageContent](pagecontent.md): una región de nivel superior en una página que contiene tipos de contenido, como Outline o Image. Un objeto PageContent se puede asignar a una posición en la página.

   - [PageContentCollection](pagecontentcollection.md): una colección de objetos PageContent que representa el contenido de una página.

- [Outline](outline.md): un contenedor para objetos Paragraph. Un Outline es un elemento secundario directo de un objeto PageContent.

- [Image](image.md): un objeto Image. Image puede ser un elemento secundario directo de un objeto PageContent o Paragraph.

- [Paragraph](paragraph.md): un contenedor para el contenido visible en una página. Un objeto Paragraph es un elemento secundario directo de un Outline.

  - [ParagraphCollection](paragraphcollection.md): Una colección de objetos Paragraph es un Outline.

- [RichText](richtext.md): un objeto RichText.

- [Table](table.md): un contenedor de objetos TableRow.

- [TableRow](tablerow.md): un contenedor de objetos TableCell.

  - [TableRowCollection](tablerowcollection.md): una colección de objetos TableRow de una tabla.
 
- [TableCell](tablecell.md): un contenedor para objetos Paragraph.

  - [TableCellCollection](tablecellcollection.md): una colección de objetos TableCell de un objeto TableRow.
        
## Recursos adicionales

- [Introducción a la programación de API de JavaScript para OneNote](../../docs/onenote/onenote-add-ins-programming-overview.md)
- [Crear el primer complemento de OneNote](../../docs/onenote/onenote-add-ins-getting-started.md)
- [Rubric Grader sample (Ejemplo de Rubric Grader)](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office Add-ins platform overview (Información general sobre la plataforma de complementos para Office)](https://dev.office.com/docs/add-ins/overview/office-add-ins)
