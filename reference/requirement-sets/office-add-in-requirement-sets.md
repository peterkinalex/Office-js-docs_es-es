# <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API de Office

Los conjuntos de requisitos son grupos de miembros de la API con nombre. Los complementos de Office usan los conjuntos de requisitos especificados en el manifiesto o usan una comprobación en tiempo de ejecución para determinar si un host de Office admite las API necesarias para el complemento. Para obtener más información, consulte [Specify Office hosts and API requirements (Especificar hosts de Office y requisitos de la API)](../docs/overview/specify-office-hosts-and-api-requirements.md).

Para obtener información sobre la compatibilidad de los complementos con el host de Office, consulte [Disponibilidad de plataformas y hosts de los complementos de Office](https://dev.office.com/add-in-availability).

## <a name="hostspecific-api-requirement-sets"></a>Conjuntos de requisitos de la API específica del host

Para obtener información sobre los conjuntos de requisitos de la API de cuadros de diálogo, de Excel, Word, OneNote y Outlook, consulte:

- [Conjuntos de requisitos de la API de JavaScript de Excel](excel-api-requirement-sets.md)
- [Conjuntos de requisitos de la API de JavaScript de Word](word-api-requirement-sets.md)
- [Conjuntos de requisitos de la API de JavaScript de OneNote](onenote-api-requirement-sets.md)
- [Información sobre los conjuntos de requisitos de la API de Outlook](../outlook/tutorial-api-requirement-sets.md)
[Conjuntos de requisitos de la API de cuadros de diálogo](dialog-api-requirement-sets.md)

## <a name="common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API

En la siguiente tabla se enumeran los conjuntos de requisitos comunes de la API, los métodos de cada conjunto, y las aplicaciones host de Office compatibles con el conjunto de requisitos. Todos estos conjuntos de requisitos de la API son de la versión 1.1.


|  Conjunto de requisitos  |  Host de Office  |  Métodos del conjunto  |
|:-----|-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint&nbsp;Online|Document.getActiveViewAsync|
| BindingEvents  | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | PowerPoint<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad<br/>Excel Online<br/>PowerPoint Online|Admite salida al formato Office Open XML (OOXML) como una matriz de bytes<br>(Office.FileType.Compressed) cuando se usa el método Document.getFileAsync.|
| CustomXmlParts    | Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DocumentEvents    | Excel<br>Excel Online<br>PowerPoint Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| Archivo  | PowerPoint<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la coerción a HTML (Office.CoercionType.Html) al leer y escribir datos mediante los métodos Document.getSelectedDataAsync,<br>Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| ImageCoercion | Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la conversión a una imagen (Office.CoercionType.Image) al escribir datos mediante el método Document.setSelectedDataAsync.|
| Buzón   |Outlook para Windows<br>Outlook para web<br>Outlook para Mac<br>Outlook Web App |Consulte [Información sobre los conjuntos de requisitos de la API de Outlook](./outlook/tutorial-api-requirement-sets.md).|
| MatrixBindings    | Excel<br>Excel Online<br>Word<br>Word Online|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la coerción a la estructura de datos de "matrix" (matriz de matrices) (Office.CoercionType.Matrix) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| OoxmlCoercion | Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la coerción al formato Open Office XML (OOXML) (Office.CoercionType.Ooxml) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| PartialTableBindings  | Access Web Apps||
| PdfFile   | PowerPoint<br/>PowerPoint Online<br/>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite salida a formato PDF (Office.FileType.Pdf)<br>al usar el método Document.getFileAsync.|
| Selección | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Configuración  | Access Web Apps<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la coerción a la estructura de datos de "table" (Office.CoercionType.Table) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| TextBindings  | Excel<br>Excel Online<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad|Admite la coerción al formato de texto (Office.CoercionType.Text) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| TextFile  | Word 2013 y posterior<br>Word 2016 para Mac<br>Word Online<br>Word para iPad<br/>|Admite salida en formato de texto (Office.FileType.Text) cuando se usa el método Document.getFileAsync.|

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



