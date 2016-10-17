
# <a name="customxmlnode-object"></a>Objeto CustomXmlNode
Representa un nodo XML en un árbol de un documento.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Modificado por última vez en**|1.1|

```js
CustomXmlNode
```


## <a name="members"></a>Miembros


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[baseName](../../reference/shared/customxmlnode.basename.md)|Obtiene el nombre base del nodo sin el prefijo de espacio de nombres, si existe alguno.|
|[nodeType](../../reference/shared/customxmlnode.nodetype.md)|Obtiene el tipo de **CustomXMLNode**.|
|[namespaceUri](../../reference/shared/customxmlnode.namespaceuri.md)|Recupera el GUID de la cadena del elemento **CustomXMLPart**.|

**Métodos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|Obtiene los nodos de forma asincrónica como una matriz de objetos **CustomXMLNode** que coinciden con la expresión XPath relativa.|
|[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|Obtiene de forma asincrónica el valor del nodo.|
|[getTextAsync](customxmlnode.gettextasync.md)|Obtiene el texto de un nodo XML de forma asincrónica en un elemento XML personalizado.|
|[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|Obtiene de forma asincrónica el XML del nodo.|
|[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|Define de forma asincrónica el valor del nodo.|
|[setTextAsync](customxmlnode.settextasync.md)|Define el texto de un nodo XML de forma asincrónica en un elemento XML personalizado.|
|[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|Define de forma asincrónica el XML del nodo.|

## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|CustomXmlParts|
|**Nivel de permisos mínimo**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word en Office para iPad.|
|1.0|Agregado|
