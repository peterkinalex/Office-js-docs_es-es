
# Objeto Bindings
Representa los enlaces que tiene el complemento en el documento.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Modificado por última vez** en|1.1|

```js
Office.context.document.bindings
```


**Propiedades**

|||
|:-----|:-----|
|Nombre|Descripción|
|[documento](../../reference/shared/bindings.document.md)|Obtiene un objeto **Document** que representa el documento asociado a este conjunto de enlaces.|

**Métodos**

|||
|:-----|:-----|
|Nombre|Descripción|
|[addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md)|Agrega un enlace a un elemento con nombre del documento.|
|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|Muestra la UI que permite al usuario especificar la selección con la que desea enlazar.|
|[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|Agrega un objeto de enlace del tipo que se ha especificado y lo enlaza con la selección actual del documento.|
|[getAllAsync](../../reference/shared/bindings.getallasync.md)|Obtiene todos los enlaces que se crearon previamente.|
|[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|Obtiene el enlace que se ha especificado por su identificador.|
|[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|Elimina el enlace especificado.|

## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||||
|:-----|:-----|:-----|:-----|
||Office para escritorio de Windows|Office Online (en el explorador)|Office para iPad|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.1|Para [addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md), [addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md) y [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md), se ha agregado compatibilidad con enlace a datos de matriz como enlace de tabla en complementos para Excel.|
|1.1|<ul><li>En cuanto a la propiedad <a href="8fa0cb4a-fad1-4f2e-9a7e-5f7aa7789eca.htm">document</a>, se ha agregado el acceso a un objeto <span class="keyword">Document</span> que representa la base de datos actual de Access en los complementos de contenido para Access.</li><li>Para todos los métodos se ha agregado compatibilidad para el enlace de tabla en los complementos de contenido para Access. </li></ul>|
|1,0|Agregado|
