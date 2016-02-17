# Referencia de la API de JavaScript de complementos de Excel

_Se aplica a: Excel 2016, Office 2016_

Los vínculos siguientes muestran los objetos de Excel de nivel avanzado disponibles en las API. Cada vínculo a la página del objeto contiene una descripción de las propiedades, las relaciones y los métodos disponibles en el objeto. Explore los vínculos siguientes para obtener más información.
	
* [Workbook](resources/workbook.md): objeto de nivel superior que contiene los objetos de libro relacionados, como hojas de cálculo, tablas, intervalos, etc. También puede usarse para enumerar las referencias relacionadas. 
* [Worksheet](resources/worksheet.md): miembro de la colección Worksheet. La colección Worksheet contiene todos los objetos de hoja de cálculo de un libro.
	* [Colección Worksheet](resources/worksheetcollection.md): colección de todos los objetos de libro que forman parte del libro. 
* [Range](resources/range.md): representa una celda, una fila, una columna o una selección de celdas con uno o más bloques contiguos de celdas.  
* [Table](resources/table.md): representa una colección de celdas organizadas diseñada para facilitar la administración de los datos. 
	* [Colección Table](resources/tablecollection.md): colección de tablas de un libro o una hoja de cálculo. 
	* [Colección TableColumn](resources/tablecolumncollection.md): colección de todas las columnas de una tabla. 
	* [Colección TableRow](resources/tablerowcollection.md): colección de todas las filas de una tabla. 
* [Chart](resources/chart.md): representa un objeto Chart de una hoja de cálculo, que es una representación visual de los datos subyacentes.   
	* [Colección Chart](resources/chartcollection.md): colección de gráficos de una hoja de cálculo.	
* [NamedItem](resources/nameditem.md): representa un nombre definido para un intervalo de celdas o un valor. Los nombres pueden ser objetos primitivos con nombre, un objeto de intervalo, etc.
	* [Colección NamedItem](resources/nameditemcollection.md): colección de objetos namedItem de un libro.
* [Binding](resources/binding.md): clase abstracta que representa un enlace a una sección del libro.
	* [Colección Binding](resources/bindingcollection.md): colección de todos los objetos de enlace que forman parte del libro. 
* [Colección TrackedObject](resources/trackedobjectscollection.md): permite que los complementos administren una referencia de objeto de intervalo en lotes sync(). 
* [RequestContext](resources/requestcontext.md): el objeto RequestContext facilita las solicitudes para la aplicación de Excel.


##### Recursos adicionales

*  [Introducción a la programación de complementos de Excel](excel-add-ins-programming-overview.md)
*  [Crear el primer complemento de Excel](build-your-first-excel-add-in.md)
*  [Explorador de fragmentos de código para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
*  [Ejemplos de código de complementos de Excel](excel-add-ins-code-samples.md) 


