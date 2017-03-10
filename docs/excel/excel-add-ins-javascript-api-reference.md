# <a name="excel-javascript-api-reference"></a>Referencia de la API de JavaScript para Excel

Puede usar la API de JavaScript para Excel para crear complementos para Excel 2016. En la lista siguiente se muestran los objetos de Excel de alto nivel que están disponibles en la API. Cada vínculo a la página del objeto contiene una descripción de las propiedades, las relaciones y los métodos disponibles en el objeto. Explore los vínculos para más información.

* [Workbook](../../reference/excel/workbook.md): objeto de nivel superior que contiene los objetos de libro relacionados, como hojas de cálculo, tablas, intervalos, etc. También puede usarse para enumerar las referencias relacionadas.
* [Worksheet](../../reference/excel/worksheet.md): miembro de la colección Worksheet. La colección Worksheet contiene todos los objetos de hoja de cálculo de un libro.
    * [Colección Worksheet](../../reference/excel/worksheetcollection.md): Colección de todos los objetos Whorksheet que forman parte del libro.
* [Range](../../reference/excel/range.md): Representa una celda, una fila, una columna o una selección de celdas con uno o más bloques contiguos de celdas.
* [Table](../../reference/excel/table.md): representa una colección de celdas organizadas diseñada para facilitar la administración de los datos.
    * [Colección Table](../../reference/excel/tablecollection.md): colección de tablas de un libro o una hoja de cálculo.
    * [Colección TableColumn](../../reference/excel/tablecolumncollection.md): colección de todas las columnas de una tabla.
    * [Colección TableRow](../../reference/excel/tablerowcollection.md): colección de todas las filas de una tabla.
* [Chart](../../reference/excel/chart.md): representa un objeto Chart de una hoja de cálculo, que es una representación visual de los datos subyacentes.
    * [Colección Chart](../../reference/excel/chartcollection.md): una colección de gráficos en una hoja de cálculo.
* [TableSort](../../reference/excel/tablesort.md): representa un objeto que ordena operaciones en objetos Table.
* [RangeSort](../../reference/excel/rangesort.md): representa un objeto que ordena operaciones en objetos Range.
* [Filter](../../reference/excel/filter.md): representa un objeto de filtro que administra el filtrado de la columna de una tabla.
* [Worksheet Protection](../../reference/excel/worksheetprotection.md): representa la protección de un objeto de hoja de cálculo.
* [Worksheet Function](../../reference/excel/functions.md): representa un contenedor de las funciones de hoja de cálculo de Microsoft Excel que pueden llamarse a través de JavaScript.
* [NamedItem](../../reference/excel/nameditem.md): representa un nombre definido para un intervalo de celdas o un valor. Los nombres pueden ser objetos con nombre primitivo, un objeto de rango, etc.
    * [Colección NamedItem](../../reference/excel/nameditemcollection.md): colección de objetos namedItem de un libro.
* [Binding](../../reference/excel/binding.md): clase abstracta que representa un enlace a una sección del libro.
    * [Colección Binding](../../reference/excel/bindingcollection.md): colección de todos los objetos de enlace que forman parte del libro.
* [Colección TrackedObject](../../reference/excel/trackedobjectscollection.md): permite que los complementos administren una referencia de objeto de intervalo en lotes sync().
* [RequestContext](../../reference/excel/requestcontext.md): el objeto RequestContext facilita las solicitudes para la aplicación de Excel.


##### <a name="additional-resources"></a>Recursos adicionales

*  [Introducción a la programación de complementos de Excel](excel-add-ins-javascript-programming-overview.md)
*  [Compilar el primer complemento de Excel](build-your-first-excel-add-in.md)
*  [Explorador de fragmentos de código para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)

