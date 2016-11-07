
# <a name="office-addin-requirement-sets"></a>Conjuntos de requisitos de complementos de Office

Los conjuntos de requisitos son grupos con nombre de miembros de la API. Los complementos de Office usan conjuntos de requisitos especificados en el manifiesto o una comprobación en tiempo de ejecución para determinar si un host de Office es compatible con las API que necesita el complemento. Para obtener más información, consulte [Especificar los hosts de Office y los requisitos de la API](../docs/overview/specify-office-hosts-and-api-requirements.md).

Para obtener una visión amplia sobre la compatibilidad de los complementos con el host de Office, consulte la página [Disponibilidad de plataformas y hosts de los complementos de Office](https://dev.office.com/add-in-availability).

## <a name="requirement-sets"></a>Conjuntos de requisitos


En la siguiente tabla se enumeran los nombres de los conjuntos de requisitos, los métodos de cada conjunto, las aplicaciones host de Office compatibles con el conjunto de requisitos y el número de versión de la API.

Para obtener información sobre los conjuntos de requisitos para Outlook, consulte [Understanding Outlook API requirement sets](./outlook/tutorial-api-requirement-sets.md) (Introducción a los conjuntos de requisitos de las API de Outlook).

|  Nombre del conjunto  |  Versión  |  Host de Office  |  Métodos del conjunto  |
|:-----|-----|:-----|:-----|
| ExcelApi   | 1.2 | Excel 2016<br>Excel Online<br>Excel para iPad<br>|Protección de hoja de cálculo<br>Funciones de hoja de cálculo<br>Ordenar<br>Filtro<br>Estilo de referencia R1C1<br>Merge Cells<br>Ajustar el alto de fila y el ancho de columna<br>Chart.getImage()<br>Range.getUsedRange(valuesOnly)|
| ExcelApi   | 1.1 | Excel 2016<br>Excel Online<br>Excel para iPad<br>|Todos los elementos del espacio de nombres de Excel.|
| WordApi    | 1.2 | Word 2016<br>Word 2016 para Mac<br>Word para iPad<br>Word Online| Todos los elementos del espacio de nombres de Word. Los métodos siguientes se han agregado a esta versión de WordApi:<br>Body.select(selectionMode)<br>Body.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>contentControl.select(selectionMode)<br>contentControl.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>inlinePicture.paragraph<br>inlinePicture.delete<br>inlinePicture.insertBreak(breakType, insertLocation)<br>inlinePicture.insertFileFromBase64(base64file, insertLocation)<br>inlinePicture.insertHtml(html, insertLocation)<br>inlinePicture.insertInlinePictureFromBase64(base64file, insertLocation)<br>inlinePicture.insertOoxml(ooxml, insertLocation)<br>inlinePicture.insertParagraph(paragraphText, insertLocation)<br>inlinePicture.insertText(text, insertLocation)<br>inlinePicture.select(selectionMode)<br>paragraph.select(selectionMode)<br>range.inlinePictures<br>range.select(selectionMode)<br>range.insertInlinePictureFomBase64(base64EcodedImage, insertLocation)|
| WordApi    | 1.1 | Word 2016<br>Word 2016 para Mac<br>Word para iPad<br>Word Online|Todos los elementos del espacio de nombres de Word, excepto los miembros de la API que se han agregado a WordApi 1.2 y versiones posteriores, que aparecen más arriba.|
| ActiveView | 1.1 | PowerPoint<br>PowerPoint Online|Document.getActiveViewAsync|
| BindingEvents  | 1.1 | Aplicaciones web de Access<br>Excel<br>Excel Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | 1.1 |PowerPoint<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad<br/>Excel Online<br/>PowerPoint Online|Admite salida al formato Office Open XML (OOXML) como una matriz de bytes<br>(Office.FileType.Compressed) cuando se usa el método Document.getFileAsync.|
| CustomXmlParts    | 1.1 |Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogAPI | 1.1 | Excel<br>PowerPoint<br>Word 2016<br>Outlook|Office.context.ui.displayDialogAsync()<br>Office.context.ui.messageParent()<br>Office.context.ui.close()|
| DocumentEvents    | 1.1 | Excel<br>Excel Online<br>PowerPoint Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| Archivo  | 1.1 | PowerPoint<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | 1.1 | Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la coerción a HTML (Office.CoercionType.Html) al leer y escribir datos mediante los métodos Document.getSelectedDataAsync,<br>Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| ImageCoercion | 1.1 | Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la conversión a una imagen (Office.CoercionType.Image) al escribir datos mediante el método Document.setSelectedDataAsync.|
| Buzón   |   | Outlook para Windows<br>Outlook para web<br>Outlook para Mac<br>Outlook Web App |consulte [Información sobre los conjuntos de requisitos de la API de Outlook](./outlook/tutorial-api-requirement-sets.md)|
| MatrixBindings    | 1.1 | Excel<br>Excel Online<br>Word<br>Word Online|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | 1.1 | Excel<br>Excel Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la coerción a la estructura de datos de "matrix" (matriz de matrices) (Office.CoercionType.Matrix) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| OoxmlCoercion | 1.1 | Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la coerción al formato Open Office XML (OOXML) (Office.CoercionType.Ooxml) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| PartialTableBindings  | 1.1 | Aplicaciones web de Access||
| PdfFile   | 1.1 | PowerPoint<br/>PowerPoint Online<br/>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite salida a formato PDF (Office.FileType.Pdf)<br>al usar el método Document.getFileAsync.|
| Selección | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Configuración  | 1.1 | Aplicaciones web de Access<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | 1.1 | Aplicaciones web de Access<br>Excel<br>Excel Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | 1.1 | Aplicaciones web de Access<br>Excel<br>Excel Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la coerción a la estructura de datos de "table" (Office.CoercionType.Table) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| TextBindings  | 1.1 | Excel<br>Excel Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la coerción al formato de texto (Office.CoercionType.Text) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| TextFile  | 1.1 | Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad<br/>|Admite salida en formato de texto (Office.FileType.Text) cuando se usa el método Document.getFileAsync.|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Métodos no incluidos en un conjunto de requisitos


Los siguientes métodos de la API de JavaScript para Office no forman parte de ningún conjunto de requisitos. Si el complemento necesita cualquiera de estos métodos, use los elementos **Methods** y **Method** del manifiesto del complemento para declarar que son necesarios o realice la comprobación en tiempo de ejecución mediante una instrucción if. Para obtener más información, consulte [Especificar los hosts de Office y los requisitos de la API](../docs/overview/specify-office-hosts-and-api-requirements.md).



|**Nombre del método**|**Compatibilidad con host de Office**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access Web App, Excel y Excel Online|
|Document.getFilePropertiesAsync|Excel, Excel Online, Word y PowerPoint|
|Document.getProjectFieldAsync|Project Standard 2013 y Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 y Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 y Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 y Project Professional 2013|
|Document.getSelectedViewAsync|PowerPoint y PowerPoint Online|
|Document.getTaskAsync|Project Standard 2013 y Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 y Project Professional 2013|
|Document.goToByIdAsync|Excel, Excel Online, Word y PowerPoint|
|Settings.addHandlerAsync|Access Web App, Excel, Excel Online, Word y PowerPoint|
|Settings.refreshAsync|Access Web App, Excel, Excel Online, Word, PowerPoint y PowerPoint Online|
|Settings.removeHandlerAsync|Access Web App, Excel, Excel Online, Word y PowerPoint|
|TableBinding.clearFormatsAsync|Excel y Excel Online|
|TableBinding.setFormatsAsync|Excel y Excel Online|
|TableBinding.setTableOptionsAsync|Excel y Excel Online|

## <a name="additional-resources"></a>Recursos adicionales



- [Especificar los hosts de Office y los requisitos de la API](../docs/overview/specify-office-hosts-and-api-requirements.md)

