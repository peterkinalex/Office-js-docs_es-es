
# <a name="understanding-the-javascript-api-for-office"></a>Información sobre la API de JavaScript para Office



Este artículo proporciona información acerca de la API de JavaScript para Office y cómo usarla. Para obtener información de referencia, consulte [API de JavaScript para Office](../../reference/javascript-api-for-office.md). Para obtener información sobre cómo actualizar los archivos de proyecto de Visual Studio a la versión más reciente de la API de JavaScript para Office, consulte [Actualizar la versión de la API de JavaScript para Office y los archivos de esquema de manifiesto](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

>**Nota:** Al generar el complemento, si va a [publicar](../publish/publish.md) el complemento en la Tienda Office, asegúrese de que se ajustan a la [directivas de validación de la Tienda Office](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Por ejemplo, para superar la validación, el complemento debe funcionar en todas las plataformas que sean compatibles con los métodos especificados en el elemento Requirements del manifiesto (vea la [sección 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3)).

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Referencia a la biblioteca de la API de JavaScript para Office en el complemento

La biblioteca de la [API de JavaScript para Office](../../reference/javascript-api-for-office.md) está formada por el archivo Office.js y los archivos .js específicos de la aplicación host asociada, como Excel-15.js y Outlook-15.js. El método más sencillo para hacer referencia a la API es usar nuestra CDN. Para hacerlo, agregue el siguiente `<script>` a la etiqueta `<head>` de la página:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Esto descargará y almacenará en caché los archivos de la API de JavaScript para Office la primera vez que se cargue el complemento para asegurarse de que usa la implementación más actualizada de Office.js y de sus archivos asociados para la versión especificada.

Para más información sobre la CDN de Office.js, incluido cómo se administra el control de versiones y la compatibilidad con versiones anteriores, vea [Referencia a la biblioteca de la API de JavaScript para Office desde su red de entrega de contenido (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Inicializar el complemento


 **Se aplica a:** todos los tipos de complementos


Office.js proporciona un evento de inicialización que se desencadena cuando la API está totalmente cargada y lista para empezar a interactuar con el usuario. Puede usar el controlador de eventos **initialize** para implementar los escenarios comunes de inicialización del complemento, como pedir al usuario que seleccione algunas celdas en Excel y después insertar un gráfico que se inicializa con los valores seleccionados. También puede usar el controlador de eventos initialize para inicializar otra lógica personalizada para el complemento, tal como establecer enlaces, pedir los valores predeterminados de la configuración del complemento, etc.

 Como mínimo, el evento initialize tendría que ser similar al ejemplo siguiente:     

```js
Office.initialize = function () { };
```
Si usa marcos de JavaScript adicionales que incluyen su propio controlador o sus propias pruebas de inicialización, tiene que colocarlos dentro del evento Office.initialize. Por ejemplo, se hará referencia a la función `$(document).ready()` de [jQuery](https://jquery.com) como se indica a continuación:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```
Es necesario que todas las páginas de un complemento de Office asignen un controlador de eventos al evento initialize, **Office.initialize**. Si no asigna un controlador de eventos, el complemento puede producir un error cuando se inicia. Además, si un usuario intenta usar el complemento con un cliente web de Office Online, como Excel Online, PowerPoint Online o Outlook Web App, este no podrá ejecutarse. Si no necesita ningún código de inicialización, el cuerpo de la función que asigne a **Office.initialize** puede estar vacío, como en el primer ejemplo anterior.

Para obtener información más detallada sobre la secuencia de eventos cuando se inicializa un complemento, consulte [Cargar el DOM y el entorno de tiempo de ejecución](../../docs/develop/loading-the-dom-and-runtime-environment.md).

#### <a name="initialization-reason"></a>Motivo de inicialización
Para los complementos de contenido y de panel, Office.initialize proporciona un parámetro _reason_ adicional. Este parámetro se puede usar para determinar cómo se agregó un complemento al documento actual. Puede usar esto para proporcionar una lógica distinta para los casos en que un complemento se inserta primero en comparación a cuando ya existe en el documento. 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
      switch (reason) {
        case 'inserted': console.log('The add-in was just inserted.');
        case 'documentOpened': console.log('The add-in is already part of the document.');
    }
}
```
Para más información, vea [Evento Office.initialize](../../reference/shared/office.initialize.md) y [Enumeración InitializationReason](../../reference/shared/initializationreason-enumeration.md) 

## <a name="context-object"></a>Objeto Context

 **Se aplica a:** todos los tipos de complementos

Cuando se inicializa un complemento, existen muchos objetos que pueden interactuar con el entorno de tiempo de ejecución. El contexto de tiempo de ejecución del complemento se ve reflejado en la API con el objeto [Context](../../reference/shared/office.context.md). **Context** es el principal objeto que proporciona acceso a los objetos más importantes de la API, como los objetos [Document](../../reference/shared/document.md) y [Mailbox](../../reference/outlook/Office.context.mailbox.md), que a su vez proporcionan acceso al contenido del documento y del buzón de correo.

Por ejemplo, en complementos de contenido o de panel de tareas, se puede usar la propiedad [document](../../reference/shared/office.context.document.md) del objeto **Context** para obtener acceso a las propiedades y los métodos del objeto **Document** e interactuar con el contenido de documentos de Word, hojas de cálculo de Excel o programaciones de Project. Del mismo modo, en los complementos de Outlook se puede usar la propiedad [mailbox](../../reference/outlook/Office.context.mailbox.md) del objeto **Context** para obtener acceso a las propiedades y los métodos del objeto **Mailbox** e interactuar con el contenido de mensajes, convocatorias de reunión o citas.

El objeto **Context** también proporciona acceso a las propiedades [contentLanguage](../../reference/shared/office.context.contentlanguage.md) y [displayLanguage](../../reference/shared/office.context.displaylanguage.md), que permiten determinar la configuración regional (el idioma) que se usa en el documento, el elemento o la aplicación host. Además, la propiedad [roamingSettings](../../reference/outlook/Office.context.md) permite tener acceso a los miembros del objeto [RoamingSettings](../../reference/outlook/RoamingSettings.md). Por último, el objeto **Context** proporciona una propiedad [ui](../../reference/shared/officeui.md) que permite al complemento iniciar cuadros de diálogo emergentes.


## <a name="document-object"></a>Objeto Document


 **Se aplica a:** tipos de complementos de panel de tareas y de contenido

Para interactuar con los datos de un documento en Excel, PowerPoint y Word, la API ofrece el objeto [Document](../../reference/shared/document.md). Puede usar miembros del objeto  **Document** para obtener acceso a los datos de las siguientes maneras:


- Leer y escribir en selecciones activas en forma de texto, celdas contiguas (matrices) o tablas.
    
- Datos tabulares (tablas o matrices).
    
- Enlaces (creados con los métodos "add" del objeto  **Bindings**).
    
- Elementos XML personalizados (solo para Word).
    
- Configuraciones o estados de complemento conservados para cada complemento en el documento.
    
También puede usar el objeto  **Document** para interactuar con datos en los documentos de Project. La funcionalidad específica de Project de la API se documenta en la clase abstracta de [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md) de los miembros. Para más información sobre cómo crear complementos de panel de tareas para Project, vea [Complementos de panel de tareas para Project](../project/project-add-ins.md).

Todas esas maneras de obtener acceso a los datos parten de una instancia del objeto  **Document** abstracto.

Puede obtener acceso a una instancia del objeto  **Document** cuando el complemento de contenido o de panel de tareas se haya inicializado con la propiedad [document](../../reference/shared/office.context.document.md) del objeto **Context**. El objeto  **Document** define las funciones comunes de acceso a datos compartidas por los documentos de Word y de Excel, y proporciona acceso al objeto **CustomXmlParts** para los documentos de Word.

El objeto  **Document** admite cuatro maneras en que los desarrolladores pueden obtener acceso al contenido de los documentos:


- Acceso basado en la selección
    
- Acceso basado en el enlace
    
- Acceso basado en elementos XML personalizados (solo en Word)
    
- Acceso basado en documentos completos (solo en PowerPoint y Word)
    
Para explicar mejor cómo funcionan los métodos de acceso a datos basados en la selección y el enlace, antes explicaremos de qué manera las API de acceso a datos proporcionan acceso a datos coherentes entre las distintas aplicaciones de Office.


### <a name="consistent-data-access-across-office-applications"></a>Acceso a datos coherentes entre las aplicaciones de Office

 **Se aplica a:** tipos de complementos de panel de tareas y de contenido

Para crear extensiones que trabajan sin ningún problema entre distintos documentos de Office, la API de JavaScript para Office extrae las particularidades de cada aplicación de Office a través de tipos de datos comunes y de la posibilidad de convertir el contenido de documentos diferentes en tres tipos de datos comunes.


#### <a name="common-data-types"></a>Tipos de datos comunes

Tanto en el acceso a datos basado en la selección como el basado en el enlace, el contenido de los documentos se expone a través de tipos de datos compartidos por todas las aplicaciones de Office compatibles. En Office 2013, hay tres tipos de datos principales compatibles:



|**Tipo de datos**|**Descripción**|**Compatibilidad con aplicación host**|
|:-----|:-----|:-----|
|Texto|Proporciona una representación de cadena de los datos de la selección o del enlace.|En Excel 2013, Project 2013 y PowerPoint 2013 solo se admite texto sin formato. En Word 2013, se admiten tres formatos de texto: texto sin formato, HTML y Office Open XML (OOXML).Cuando se selecciona texto en una celda de Excel, los métodos basados en selecciones leen y escriben en todo el contenido de la celda, aunque solo esté seleccionada una parte del texto en la celda. Cuando se selecciona texto en Word y PowerPoint, los métodos basados en selecciones leen y escriben solo la secuencia de caracteres seleccionada.Project 2013 y PowerPoint 2013 solo admiten el acceso a datos basado en la selección.|
|Matriz|Proporciona los datos de la selección o del enlace como una  **Array** bidimensional, implementada en JavaScript como una matriz de matrices.Por ejemplo, dos líneas de valores de  **string** en dos columnas serían ` [['a', 'b'], ['c', 'd']]` y una sola columna de tres filas sería `[['a'], ['b'], ['c']]`.|El acceso a los datos de la matriz solo se admite en Excel 2013 y Word 2013.|
|Tabla|Proporciona los datos de la selección o del enlace como un objeto [TableData](../../reference/shared/tabledata.md). El objeto  **TableData** expone los datos a través de las propiedades **headers** y **rows**.|El acceso a los datos de la tabla solo se admite en Excel 2013 y Word 2013.|

#### <a name="data-type-coercion"></a>Coerción de tipos de datos

Los métodos de acceso a datos en los objetos  **Document** y [Binding](../../reference/shared/binding.md) admiten la especificación del tipo de datos deseado mediante el uso del parámetro _coercionType_ de dichos métodos y los correspondientes valores de enumeración [CoercionType](../../reference/shared/coerciontype-enumeration.md). Independientemente de la forma real del enlace, las distintas aplicaciones de Office admiten los tipos de datos comunes tal intentar convertir los datos en el tipo de datos solicitado. Por ejemplo, si se selecciona un párrafo o una tabla de Word, el desarrollador puede especificar que se lea como texto sin formato, como texto HTML, Office Open XML o como una tabla; por su parte, la implementación de la API administra las transformaciones y las conversiones de datos necesarias.


 >**Sugerencia**   **¿Cuándo debería usar la matriz en vez de la tabla coercionType para el acceso a los datos?** Si necesita que los datos tabulares crezcan de forma dinámica cuando se agreguen filas y columnas, y debe trabajar con los encabezados de la tabla, debe usar el tipo de datos de tabla (para ello, especifique el parámetro _coercionType_ de un método de acceso de datos del objeto **Document** o **Binding** como `"table"` o **Office.CoercionType.Table**). La adición de filas y columnas en la estructura de datos se admite tanto en los datos de matriz como de tabla, pero la anexión de filas y columnas solo se admite para los datos de tabla. Si no planea agregar filas y columnas, y los datos no requieren la función de encabezados, entonces debe usar el tipo de datos de matriz (para ello, especifique el parámetro  _coercionType_ del método de acceso a los datos como `"matrix"` o **Office.CoercionType.Matrix**), que le proporcionará un modelo más sencillo de interacción con los datos.

Si los datos no se pueden convertir a un tipo especificado, la propiedad [AsyncResult.status](../../reference/shared/asyncresult.error.md) en la devolución de llamada devuelve `"failed"`. En ese caso, se puede usar la propiedad [AsyncResult.error](../../reference/shared/asyncresult.context.md) para obtener acceso al objeto [Error](../../reference/shared/error.md) con información sobre por qué falló la llamada del método.


## <a name="working-with-selections-using-the-document-object"></a>Trabajar con selecciones con el objeto Document


El objeto  **Document** expone métodos que le permiten leer y escribir en la selección actual del usuario con el modo "obtener y olvidar". Para hacerlo, el objeto **Document** proporciona los métodos **getSelectedDataAsync** y **setSelectedDataAsync**.

Para ver ejemplos de código que muestran cómo realizar tareas con selecciones, vea [Leer y escribir datos en la selección activa de un documento u hoja de cálculo](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>Trabajar con enlaces con los objetos Bindings y Binding


El acceso a datos basado en el enlace permite a los complementos de panel de tareas y de contenido obtener acceso sistemáticamente a una determinada región de un documento u hoja de cálculo a través de un identificador asociado con un enlace. Primero, la aplicación debe establecer el enlace. Para hacerlo, llama a uno de los métodos que asocian una parte del documento con un identificador único: [addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md), [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) o [addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md). Una vez se ha establecido el enlace, el complemento puede usar el identificador para acceder a los datos de la región asociada del documento o la hoja de cálculo. La creación de enlaces aporta al complemento las siguientes ventajas:


- Permite el acceso a estructuras de datos comunes de aplicaciones de Office compatibles como, por ejemplo, tablas, rangos o texto (secuencia de caracteres contiguos).
    
- Permite las operaciones de lectura/escritura sin que el usuario tenga que hacer ninguna selección.
    
- Establece una relación entre el complemento y los datos del documento. Los enlaces se conservan en el documento y es posible tener acceso a estos más adelante.
    
Al establecer un enlace, también puede subscribirse a eventos de cambio de datos y de selección designados en esa región en concreto del documento u hoja de cálculo. Es decir, que al complemento solo se le notifican los cambios que ocurren dentro de la región delimitada, y no los cambios generales que se den en todo el documento u hoja de cálculo.

El objeto [Bindings](../../reference/shared/bindings.bindings.md) expone un método [getAllAsync](../../reference/shared/bindings.getallasync.md) que permite el acceso al conjunto de todos los enlaces establecidos en el documento u hoja de cálculo. Se puede tener acceso a enlaces individuales por el id. a través de los métodos [Bindings.getBindingByIdAsync](../../reference/shared/bindings.getbyidasync.md) o [Office.select](../../reference/shared/office.select.md). Se pueden establecer nuevos enlaces, así como quitar los enlaces existentes, con uno de los siguientes métodos del objeto **Bindings**: [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md), [addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md), [addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md) o [releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md).

Existen tres tipos de enlaces distintos que puede especificar con el parámetro  _bindingType_ al crear un enlace con los métodos **addFromSelectionAsync**, **addFromPromptAsync** o **addFromNamedItemAsync**:



|**Tipo de enlace**|**Descripción**|**Compatibilidad con aplicación host**|
|:-----|:-----|:-----|
|Enlace de texto|Enlaza con una región del documento que se puede representar como texto.|En Word, la mayoría de las selecciones contiguas son válidas, mientras que en Excel solo se puede seleccionar una celda como destino del enlace de texto. En Excel, solo se admite el texto sin formato. En Word pueden usarse tres formatos: texto sin formato, HTML y Open XML para Office.|
|Enlace de matriz|Enlaza con una región fija de un documento que contiene datos tabulares sin encabezados.Los datos de los enlaces de matriz se escriben o se leen como una  **Array** de dos dimensiones que en JavaScript se implementa como una matriz de matrices. Por ejemplo, dos filas de valores de **string** de dos columnas pueden escribirse o leerse como ` [['a', 'b'], ['c', 'd']]`. Asimismo, una única columna de tres filas puede escribirse o leerse como  `[['a'], ['b'], ['c']]`.|En Excel, se puede usar cualquier selección contigua de celdas para establecer un enlace de matriz. En Word, solo las tablas admiten enlaces de matriz.|
|Enlace de tabla|Enlaza con una región de un documento que contiene una tabla con encabezados. Los datos de un enlace de tablas se escriben o leen como un objeto [TableData](../../reference/shared/tabledata.md). El objeto **TableData** expone los datos a través de las propiedades **headers** y **rows**.|Se puede usar como base cualquier tabla de Excel o de Word para establecer un enlace de tabla. Una vez se haya establecido un enlace de tabla, cada fila o columna nueva que se agregue a la tabla se incluirá automáticamente al enlace. |
Cuando haya creado un enlace con uno de los tres métodos "add" del objeto  **Bindings**, puede trabajar con los datos y las propiedades del enlace mediante el uso de los métodos del objeto correspondiente: [MatrixBinding](../../reference/shared/binding.matrixbinding.md), [TableBinding](../../reference/shared/binding.tablebinding.md) o [TextBinding](../../reference/shared/binding.textbinding.md). Estos tres objetos heredan los métodos [getDataAsync](../../reference/shared/binding.getdataasync.md) y [setDataAsync](../../reference/shared/binding.setdataasync.md) del objeto **Binding** que le permiten interactuar con los datos enlazados.

Para ver ejemplos de código que muestran cómo realizar tareas con enlaces, vea [Enlazar a regiones de un documento u hoja de cálculo](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>Trabajar con elementos XML personalizados con los objetos CustomXmlParts y CustomXmlPart


 **Se aplica a:** complementos de panel de tareas para Word

Los objetos [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md) y [CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md) de la API proporcionan acceso a elementos XML personalizados en documentos de Word, que permiten la manipulación controlada por XML de los contenidos del documento. Para ver una demostración de cómo se trabaja con los objetos **CustomXmlParts** y **CustomXmlPart**, vea el ejemplo de código [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts).


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>Trabajar con todo el documento con el método getFileAsync


 **Se aplica a:** complementos de panel de tareas para Word y PowerPoint

El método [Document.getFileAsync](../../reference/shared/document.getfileasync.md) y los miembros de los objetos [File](../../reference/shared/file.md) y [Slice](../../reference/shared/slice.md) proporcionan funciones para obtener archivos de documentos completos de Word y PowerPoint en segmentos (fragmentos) de hasta 4 MB a la vez. Para más información, vea el tema sobre [cómo obtener todo el contenido de archivos de un documento en un complemento](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md).


## <a name="mailbox-object"></a>Objeto Mailbox


 **Se aplica a:** complementos de Outlook

Los complementos de Outlook usan, principalmente, un subconjunto de la API que se expone a través del objeto [Mailbox](../../reference/outlook/Office.context.mailbox.md). Para obtener acceso a los objetos y miembros que se usan específicamente en los complementos de Outlook, como el objeto [Item](../../reference/outlook/Office.context.mailbox.item.md), use la propiedad [mailbox](../../reference/outlook/Office.context.mailbox.md) del objeto **Context** para obtener acceso al objeto **Mailbox**, como se muestra en la línea de código siguiente.




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Además, los complementos de Outlook pueden usar los objetos siguientes:


-  Objeto de **Office**: para inicialización.
    
-  Objeto **Context**: para obtener acceso al contenido y para mostrar las propiedades de idioma.
    
-  Objeto **RoamingSettings**: para guardar configuraciones personalizadas específicas del complemento de Outlook en el buzón del usuario donde está instalado el complemento.
    
Para obtener información sobre el uso de JavaScript en los complementos de Outlook, vea [Complementos de Outlook](../outlook/outlook-add-ins.md) y [Introducción a las características y la arquitectura de los complementos de Outlook](../outlook/overview.md).


## <a name="api-support-matrix"></a>Matriz de compatibilidad de la API


En esta tabla se resumen la API y las características compatibles con los distintos tipos de complementos (contenido, panel de tareas y Outlook), así como las aplicaciones de Office que pueden hospedarlas cuando se especifican las [aplicaciones host de Office compatibles con el complemento](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx) usando el [esquema de manifiesto de la aplicación 1.1 y las características compatibles con la API de JavaScript v1.1 para Office](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
||**Nombre de host**|Base de datos|Libro|Buzón|Presentación|Documento|Project|
||**Aplicaciones host** **compatibles**|aplicaciones web de Access|ExcelExcel Online|OutlookOutlook Web AppOWA para dispositivos|PowerPointPowerPoint Online|Word|Project|
|**Tipos de complementos admitidos**|Contenido|v|v||v|||
||Panel de tareas||v||v|v|v|
||Outlook|||v||||
|**Características de API compatibles**|Leer/escribir texto||v||v|v|v (solo lectura)|
||Leer/escribir matriz||v|||v||
||Leer/escribir tabla||v|||v||
||Leer/escribir HTML|||||v||
||Leer/escribir Office Open XML|||||v||
||Leer propiedades de tareas, recursos, vistas y campos||||||v|
||Eventos de cambio de selección||v|||v||
||Obtener el documento completo||||v|v||
||Enlacesy eventos de enlace|v (solo enlaces totales y parciales de tabla)|v|||v||
||Leer/escribir elementos XML personalizados|||||v||
||Datos persistentes de estado del complemento(configuración)|v (por complemento de host)|v (por documento)|v (por buzón)|v (por documento)|v (por documento)||
||Eventos de cambio de configuración|v|v||v|v||
||Obtener eventos de modo de vista activay cambio de vista||||v|||
||Desplazarse a ubicacionesdel documento||v||v|v||
||Activar de forma contextualcon reglas y regex|||v||||
||Leer propiedades del elemento|||v||||
||Leer perfil del usuario|||v||||
||Obtener datos adjuntos|||v||||
||Obtener token de identidad del usuario|||v||||
||Llamar a los servicios Web Exchange|||v||||
