
# Objeto CustomXmlPart
Representa un único objeto **CustomXMLPart** de una colección de objetos [CustomXMLParts](../../reference/shared/customxmlparts.customxmlparts.md).

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Modificado por última vez en**|1.1|

```
Office.context.document.customXmlParts.getByIdAsync(id);
```


## Miembros


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[builtIn](../../reference/shared/customxmlpart.builtin.md)|Obtiene un valor que indica si el objeto CustomXMLPart se encuentra integrado.|
|[id](../../reference/shared/customxmlpart.id.md)|Obtiene el GUID del elemento XML personalizado.|
|[namespaceManager](../../reference/shared/customxmlpart.namespacemanager.md)|Obtiene el conjunto de asignaciones de prefijo de espacio de nombres (CustomXMLPrefixMappings) que se usa en el objeto CustomXMLPart actual.|

**Métodos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md)|Agrega de forma asincrónica un controlador de eventos para un evento del objeto **CustomXmlPart**.|
|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|Elimina de forma asincrónica este elemento XML personalizado de la colección.|
|[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|Obtiene de forma asincrónica cualquier objeto CustomXmlNodes de un elemento XML personalizado que coincide con la expresión XPath especificada.|
|[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|Obtiene de forma asincrónica el contenido XML de un elemento XML personalizado.|
|[removeHandlerAsync](../../reference/shared/customxmlpart.removehandlerasync.md)|Quita un controlador de eventos para un evento del objeto **CustomXmlPart**.|

**Eventos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md)|Ocurre cuando se suprime un nodo.|
|[dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md)|Ocurre cuando se inserta un nodo.|
|[dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md)|Se genera al reemplazar un nodo.|

## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|CustomXmlParts|
|**Nivel de permisos mínimo**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word en Office para iPad.|
|1.0|Agregado|
