
# Dar formato a tablas en complementos para Excel


En este artículo se explican las diferentes características de la API de formato y se describe cómo usarlas. En esta versión, puede especificar mediante programación el formato de celda y otras opciones solo para tablas (no para las estructuras de datos  **Office.CoercionType.Text** o **Office.CoercionType.Matrix**), y solo en complementos de Excel. Para establecer el formato en un complemento:

- El usuario selecciona la tabla (o el lugar donde desea insertar la tabla mediante programación) y, a continuación, el complemento puede llamar al método  **Document.setSelectedDataAsync** en dicha tabla para establecer el formato.

- Si el libro ya contiene tablas enlazadas (o el complemento usa uno de los métodos "addFrom" del objeto [Bindings](../../reference/shared/bindings.bindings.md) para crear tablas enlazadas cuando se inicializa), su complemento puede llamar al método **Binding.setDataAsync** en esas tablas enlazadas para establecer el formato.
    
>**Importante:** Para usar estos métodos nuevos y actualizados con el fin de aplicar formato a tablas en complementos de Excel, el proyecto del complemento debe [usar o estar actualizado para usar Office.js v1.1 o posterior](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

## Especificar formato

Para especificar el formato que desea establecer, cree un literal de objeto JavaScript que contenga uno o varios pares clave-valor. Si lo desea, puede combinar una serie de valores de formato en una lista dentro del objeto JavaScript. Por ejemplo: 


```js
var myFormat = {fontStyle:"bold", width:"autoFit", borderColor:"purple"};
```

Para aplicar el formato, pase el objeto JavaScript a uno de los métodos compatibles con la aplicación de formato a datos y otras características de la tabla.

Puede trabajar con el formato de dos formas:


- La primera vez que el complemento escriba datos en una selección o un enlace, especifique los parámetros opcionales  _cellFormat_ o _tableOptions_ en el objeto _options_ pasado a los métodos [Document.setSelectedDataAysnc](../../reference/shared/document.setselecteddataasync.md) o [Binding.setDataAsync](../../reference/shared/binding.setdataasync.md).
    
- Tras establecer el formato inicial, puede [borrarlo o actualizarlo](#borrarlo-o-actualizarlo) con uno de los nuevos métodos dedicados a este fin.
    

## Uso de parámetros opcionales con métodos de configuración de datos

En cuanto a los enlaces de tabla, si decide establecer los datos con los métodos  **Document.setSelectedData** o **Binding.setDataAsync**, puede especificar el formato con los parámetros opcionales  _tableOptions_ y _cellFormat_.


### El parámetro opcional tableOptions

Use el parámetro opcional  _tableOptions_ para especificar los estilos de tabla predeterminados y activar o desactivar ciertas características de tabla (como **Fila de encabezado**,  **Fila Total** y **Filas con bandas**). El valor que pase como parámetro  _tableOptions_ deberá ser un objeto JavaScript que contenga una lista de pares clave-valor. Por ejemplo:


```js
tableOptions: {bandedRows: true, filterButton: false, style:"TableStyleMedium3"};
```


### El parámetro opcional cellFormat

Use el parámetro opcional  _cellFormat_ para cambiar los valores de formato de celda (como el ancho, el alto, la fuente, el fondo, la alineación, etc.). El valor que pase como parámetro _cellFormat_ deberá ser una matriz que contenga una lista de los objetos JavaScript que especifiquen las celdas que deben considerarse celdas de destino y los formatos que se les deben aplicar. Por ejemplo:


```js
cellFormat: 
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: Office.Table.Headers, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}]
```

Puede combinar varios pares  `cells:` y `format:` en la matriz _cellFormat_ para minimizar el número de llamadas de función necesarias para aplicar el formato.


#### cells

Use  `cells:` para especificar el rango de columnas, filas y celdas a las que desea aplicar formato.


**Rangos compatibles con los valores de las celdas**


|**configuración del rango de celdas**|**Descripción**|
|:-----|:-----|
| `{row: i}`|Especifica el rango que se extiende hasta la fila ith de datos de la tabla.|
| `{column: i}`|Especifica el rango que se extiende hasta la columna ith de datos de la tabla.|
| `{row: i, column: j}`|Especifica el rango de celdas desde la fila ith hasta la columna jth de datos de la tabla.|
| `Office.Table.All`|Especifica toda la tabla, incluidos los encabezados de columna, los datos y los totales (si resulta aplicable).|
| `Office.Table.Data`|Especifica solo los datos de la tabla (no los encabezados ni los totales).|
| `Office.Table.Headers`|Especifica solo la fila de encabezado.|

#### format

Use  `format:` para especificar el formato que quiere aplicar al rango definido con `cells:` como una lista de pares clave-valor de JavaScript. Para ver una lista de los valores posibles, consulte [Claves y valores de formato compatibles](#claves-y-valores-de-formato-compatibles).

 **Limita la especificación de formato para Excel Online**

Al establecer formato en Excel Online, el número de  _grupos de formato_ que se pasa al parámetro _cellFormat_ no puede ser superior a 100. Un único grupo de formato consta de un conjunto de formato aplicado a un rango de celdas especificado. (Es decir, todo lo especificado en uno de los literales de objeto `cells:` en la matriz pasada a _cellFormat_.) Por ejemplo, la siguiente llamada pasa dos grupos de formato a  _cellFormat_.




```js
Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```


#### Aplicar parámetros opcionales

En esta versión, solo los métodos  **Document.setSelectedDataAsync** y **TableBinding.setDataAsync** permiten escribir datos y establecer el formato de las tablas en la misma llamada mediante los parámetros opcionales _tableOptions_ y _cellFormat_. En los siguientes ejemplos, el valor de  `tableData` pasado al primer parámetro de cada método (el parámetro _data_) debe ser un objeto [TableData](../../reference/shared/tabledata.md) que contenga la definición de la tabla y los datos que se van a escribir.

 **Ejemplo con Document.setSelectedDataAsync**




```js
Office.context.document.setSelectedDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 **Ejemplo con TableBinding.setDataAsync**




```js
Office.select("bindings#myBinding").setDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 >**Nota:**: En la llamada a `Office.select("bindings#myBinding")` se supone que ya existe un enlace denominado `myBinding` en la hoja de cálculo.


## Actualizar y borrar el formato


Cuando se establece formato con los parámetros opcionales  _cellFormat_ y _tableOptions_ de los métodos **Document.setSelectedDataAsync** o **TableBinding.setDataAsync**, solo se define el formato la primera vez que se llama a estos métodos. Para actualizar o borrar el formato, debe usar tres métodos nuevos del objeto  **TableBinding**:  **setFormatsAsync**,  **setTableOptionsAsync** y **clearFormatsAsync**.


### Actualizar el formato

El método [TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md) solo se emplea para actualizar el formato de celda (como el ancho, el alto, la fuente, el fondo y la alineación). Usa _cellFormat_ como parámetro necesario:


```js
Office.select("bindings#myBinding").setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

El método [TableBinding.setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md) solo se emplea para actualizar las opciones de tabla (como las bandas en filas y los botones de filtro). Usa _tableOptions_ como parámetro necesario:




```js
var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
```


### Borrar el formato

El método [TableBinding.clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md) se emplea para borrar todo el formato de la tabla. Usa el parámetro opcional _asyncContext_ y una función de devolución de llamada opcional:


```js
Office.select("bindings#myBinding").clearFormatsAsync();
```


## Claves y valores de formato compatibles


En las siguientes tablas se muestran los pares clave-valor posibles que se pueden pasar a los parámetros  _cellFormat_ o _tableOptions_.

Para los valores de  `format:`, las opciones disponibles se corresponden con un conjunto de las opciones del cuadro de diálogo  **Formato de celdas** (haga clic con el botón secundario en > **Formato de celdas** o en **Formato** > **Formato de celdas** en la pestaña **Inicio** de la cinta). Para los valores de `tableOptions:`, las opciones se correspondes con las de los grupos  **Opciones de estilo de tabla** y **Estilos de tabla** de la pestaña **Herramientas de tabla** |**Diseño** de la cinta.


 >**Importante**:  Los métodos de la API de formato solo admiten las opciones y valores que se recogen a continuación. Si especifica opciones o valores de formato distintos de los mencionados, no está definido el comportamiento de control. Estos comportamientos de control sin definir no son necesariamente coherentes en las plataformas compatibles; no debe desarrollar los complementos según ninguno de los efectos secundarios de estos comportamientos no definidos para cualquier plataforma específica. En cambio, los comportamientos de control no definidos no deben dañar el estado y la interfaz de usuario del complemento o los documentos con los que interactúan.


**Alineación**


|**Tecla**|**Valores**|**Notas**|
|:-----|:-----|:-----|
| `alignHorizontal:`|"general" \| "left" \| "center" \| "right" \| "fill" \| "justify" \| "center across selection" \| "distributed"|Cuando se combina con un valor de sangría, solo se admiten las siguientes combinaciones:<br/><br/><ul><li><code>alignHorizontal: "left"</code> y <code>indentLeft: \<value\></code></li></ul><ul><li><code>alignHorizontal: "right"</code> y <code>indentRight: \<value\></code></li></ul><ul><li><code>alignHorizontal: "distributed"</code> y <code>indentDistributed: \<value\></code></li></ul>|
| `alignVertical:`|"top" \| "center" \| "bottom" \| "justify" \| "distributed"||



**Fondo**


|**Tecla**|**Valores**|**Notas**|
|:-----|:-----|:-----|
| `backgroundColor:`|"none" \| \<Todos los nombres de colores predefinidos\> \| #RRGGBB|Nombres de colores predefinidos:<br/><br/>"black", "blue", "gray", "green", "orange", "pink", "purple", "red", "teal", "turquoise", "violet", "white", "yellow"|



**Borde**


|**Tecla**|**Valores**|**Notas**|
|:-----|:-----|:-----|
| `borderStyle:`|"none" \| \<Todos los nombres de estilo de borde predefinidos\>|Nombres de estilo de borde predefinidos<br/><br/>"dash dot", "dash dot dot", "dashed", "dotted", "double", "hair", "medium dash dot", "medium dash dot dot", "medium dashed", "medium", "slant dash dot", "thick", "thin"<br/><br/>Se aplica a todos los bordes del rango especificado. (Equivale a especificar estilos de borde con los valores predefinidos **Contorno** e **Interior** en la pestaña **Borde** del cuadro de diálogo **Formato de celdas**).<br/><br/> **Nota:** Excel 2013 puede representar los 13 estilos de borde predefinidos. En cambio, Excel Online no es compatible con todos los estilos de borde. La tabla siguiente describe la representación que se usa para cada estilo de borde al abrir la hoja de cálculo en Excel Online.<br/><br/><table><tr><th>Excel 2013</th><th>Excel Online</th></tr><tr><td>"dash dot"</td><td>dashed (1px)</td></tr><tr><td>"dash dot dot"</td><td>dotted (1px)</td></tr><tr><td>"dashed"</td><td>dotted (1px)</td></tr><tr><td>"dotted"</td><td>dashed (1px)</td></tr><tr><td>"double"</td><td>double (3px)</td></tr><tr><td>"hair"</td><td>solid (1px)</td></tr><tr><td>"medium dash dot"</td><td>dashed (2px)</td></tr><tr><td>"medium dash dot dot"</td><td>dotted (2px)</td></tr><tr><td>"medium dashed"</td><td>dashed (2px)</td></tr><tr><td>"medium"</td><td>solid (2px)</td></tr><tr><td>"slant dash dot"</td><td>dashed (2px)</td></tr><tr><td>"thick"</td><td>solid (3px)</td></tr><tr><td>"thin"</td><td>solid (1px)</td></tr></table>|
| `borderColor:`|"automatic" \| \<Todos los nombres de colores predefinidos\> \| #RRGGBB|Se aplica a todos los bordes del rango especificado.|
| `borderTopStyle:`|"none" \| \<Todos los nombres de estilo de borde predefinidos\>|Se aplica a todos los bordes del rango especificado.|
| `borderTopColor:`|"automatic" \| \<Todos los nombres de colores predefinidos\> \| #RRGGBB|Se aplica a todos los bordes del rango especificado.|
| `borderBottomStyle:`|"none" \| \<Todos los nombres de estilo de borde predefinidos\>|Se aplica a todos los bordes del rango especificado.|
| `borderBottomColor:`|"automatic" \| \<Todos los nombres de colores predefinidos\> \| #RRGGBB|Se aplica a todos los bordes del rango especificado.|
| `borderLeftStyle:`|"none" \| \<Todos los nombres de estilo de borde predefinidos\>|Se aplica a todos los bordes del rango especificado.|
| `borderLeftColor:`|"automatic" \| \<Todos los nombres de colores predefinidos\> \| #RRGGBB|Se aplica a todos los bordes del rango especificado.|
| `borderRightStyle:`|"none" \| \<Todos los nombres de estilo de borde predefinidos\>|Se aplica a todos los bordes del rango especificado.|
| `borderRightColor:`|"automatic" \| \<Todos los nombres de colores predefinidos\> \| #RRGGBB|Se aplica a todos los bordes del rango especificado.|
| `borderOutlineStyle:`|"none" \| \<Todos los nombres de estilo de borde predefinidos\>|Se aplica a todos los bordes del rango especificado.|
| `borderOutlineColor:`|"automatic" \| \<Todos los nombres de colores predefinidos\> \| #RRGGBB|Se aplica a todos los bordes del rango especificado.|
| `borderInlineStyle:`|"none" \| \<Todos los nombres de estilo de borde predefinidos\>|Solo se aplica a los bordes interiores del rango especificado. Equivale a especificar los estilos de borde únicamente con el valor preestablecido  **Interior** de la pestaña **Borde** del cuadro de diálogo **Formato de celdas**.|
| `borderInlineColor:`|"automatic" \| \<Todos los nombres de colores predefinidos\> \| #RRGGBB|Solo se aplica a los bordes interiores del rango especificado. |



**Ajuste, alto y ancho de celda**


|**Tecla**|**Valores**|
|:-----|:-----|
| `width:`|"auto fit" \|  **Número**|
| `height:`|"auto fit" \|  **Número**|
| `wrapping:`|**Boolean**|



**Fuente**


|**Tecla**|**Valores**|**Notas**|
|:-----|:-----|:-----|
| `fontFamily:`|\<Todos los nombres de fuente disponibles\>|Cuando establece una fuente en Excel Online, si la fuente no está disponible en el explorador, la API intentará regresar a las siguientes fuentes en este orden: Segoe UI, Thonburi, Arial, Verdana y Microsoft Sans Serif. Si ninguna de estas fuentes está disponible, se usa la fuente predeterminada del explorador.|
| `fontStyle:`|"regular" \| "italic" \| "bold" \| "bold italic"|**Nota**: En el momento de esta publicación, el proceso de establecer `fontStyle:` en "italic" y, después, "bold" (o viceversa) se comporta como una unión de estos dos valores. Es decir, si, por ejemplo, primero establece "italic" y más tarde establece "bold", el resultado será "bold italic". Para establecer bold o italic _solo_ en un rango configurado anteriormente en bold o italic, primero debe establecer `fontStyle:"regular"` para borrar el formato anterior.|
| `fontSize:`|**Número**||
| `fontUnderlineStyle:`|"none" \| "single" \| "double" \| "single accounting" \| "double accounting"||
| `fontColor:`|"automatic" \| \<Todos los nombres de colores predefinidos\> \| #RRGGBB||
| `fontDirection:`|"context" \| "left-to-right" \| "right-to-left"|Excel Online no admite en este momento la visualización de texto de derecha a izquierda. En cambio, si el complemento establece  `fontDirection:` "de derecha a izquierda" cuando se ejecuta en Excel Online, esa configuración de formato se guarda en el archivo del libro y se muestra correctamente al abrir el libro en Excel de escritorio.|
| `fontStrikethrough:`|**Booleano**||
| `fontSuperscript:`|**Booleano**||
| `fontSubScript:`|**Booleano**||
| `fontNormal:`|**Booleano**|Establece la fuente, el estilo de fuente, el tamaño y los efectos en el estilo normal. Esto restablece el formato de fuente de celda a los valores predeterminados. Equivale a seleccionar la casilla **Fuente normal** en la pestaña **Fuente** del  cuadro de diálogo **Formato de celdas**.|



**Sangría**


|**Tecla**|**Valores**|**Notas**|
|:-----|:-----|:-----|
| `indentLeft:`|**Número**|Cuando se combina con un valor de alineación, se admiten solo las siguientes combinaciones:<br/><br/><ul><li><code>alignHorizontal: "left"</code> y <code>indentLeft: \<value\></code></li></ul>|
| `indentRight:`|**Número**|Cuando se combina con un valor de alineación, se admiten solo las siguientes combinaciones:<br/><br/><ul><li><code>alignHorizontal: "right"</code> y <code>indentRight: \<value\></code></li></ul>|
| `indentDistributed:`|**Número**|Cuando se combina con un valor de alineación, se admiten solo las siguientes combinaciones:<br/><br/><ul><li><code>alignHorizontal: "distributed"</code> y <code>indentDistributed: \<value\></code></li></ul>|



**Formato de los números**


|**Tecla**|**Valores**|**Notas**|
|:-----|:-----|:-----|
| `numberFormat:`|**Cadena**|Para especificar el formato de número, use una cadena de formato de número personalizado. Por ejemplo, para especificar dos posiciones decimales con una coma como separador de miles, debería especificar:<br/><br/> `numberFormat:"#,###.00"`<br/><br/>Estas son las mismas cadenas de formato personalizado que puede [crear con la categoría Formato personalizado en la pestaña Número del cuadro de diálogo Formato de celdas](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1).<br/><br/> **Sugerencia:** Puede ver el aspecto de una cadena de formato para una categoría estándar en el cuadro de diálogo **Formato de celdas** en Excel con los siguientes pasos:<br/><br/><ol><li>Seleccione una categoría de formato estándar, por ejemplo <span class="ui">Moneda</span>, de la lista <b>Categoría</b>.</li><li>Establezca las opciones de formato en el lado derecho del cuadro de diálogo.</li><li>Seleccione la categoría <b>Personalizado</b> para ver la cadena de formato en la parte superior de la lista <b>Tipo</b>.</li></ol>|



**Opciones de tabla**


|**Tecla**|**Valores**|**Notas**|
|:-----|:-----|:-----|
| `style:`|"none" \| \<Todos los nombres de estilo de tabla predefinidos\>|Nombres de estilo de tabla predefinidos:<br/><br/>"TableStyleLight1", "TableStyleLight2", "TableStyleLight3", "TableStyleLight4", "TableStyleLight5", "TableStyleLight6", "TableStyleLight7", "TableStyleLight8", "TableStyleLight9", "TableStyleLight10", "TableStyleLight11", "TableStyleLight12", "TableStyleLight13", "TableStyleLight14", "TableStyleLight15", "TableStyleLight16", "TableStyleLight17", "TableStyleLight18", "TableStyleLight19", "TableStyleLight20", "TableStyleLight21", "TableStyleMedium1", "TableStyleMedium2", "TableStyleMedium3", "TableStyleMedium4", "TableStyleMedium5", "TableStyleMedium6", "TableStyleMedium7", "TableStyleMedium8", "TableStyleMedium9", "TableStyleMedium10", "TableStyleMedium11", "TableStyleMedium12", "TableStyleMedium13", "TableStyleMedium14", "TableStyleMedium15", "TableStyleMedium16", "TableStyleMedium17", "TableStyleMedium18", "TableStyleMedium19", "TableStyleMedium20", "TableStyleMedium21", "TableStyleMedium22", "TableStyleMedium23", "TableStyleMedium24", "TableStyleMedium25", "TableStyleMedium26", "TableStyleMedium27", "TableStyleMedium28", "TableStyleDark1", "TableStyleDark2", "TableStyleDark3", "TableStyleDark4", "TableStyleDark5", "TableStyleDark6", "TableStyleDark7", "TableStyleDark8", "TableStyleDark9", "TableStyleDark10", "TableStyleDark11"<br/><br/>Para ver cómo es un estilo de tabla, inserte una tabla en Excel, en las **Herramientas de tabla** \| **Diseño**, elija la lista desplegable  **Estilos rápidos** y seleccione un estilo predefinido. La información sobre herramientas para el estilo corresponderá a uno de los valores de la lista anterior.|
| `headerRow:`|**Boolean**||
| `firstColumn:`|**Booleano**||
| `filterButton:`|**Booleano**||
| `totalRow:`|**Booleano**||
| `lastColumn:`|**Booleano**||
| `bandedRows:`|**Booleano**||
| `bandedColumns:`|**Boolean**||
