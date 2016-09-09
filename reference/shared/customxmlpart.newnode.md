
# Propiedad NodeInsertedEventArgs.newNode
Obtiene el nodo que se acaba de agregar al objeto **CustomXMLPart**.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Modificado por última vez en**|1.1|

```
var myNode = eventArgsObj.newNode;
```


## Valor devuelto

Un objeto [CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md) que representa el nodo que se acaba de insertar.


## Comentarios

Tenga en cuenta que este nodo puede tener elementos secundarios si se acaba de agregar un subárbol al documento.


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

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
