
# <a name="nodedeletedeventargs.oldnode-property"></a>Propiedad NodeDeletedEventArgs.oldNode
Obtiene el nodo que se acaba de eliminar del objeto **CustomXmlPart**.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Modificado por última vez en**|1.1|

```
var myNode = eventArgsObj.oldNode;
```


## <a name="return-value"></a>Valor devuelto

Un objeto [CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md) que representa el nodo que se acaba de eliminar.


## <a name="remarks"></a>Comentarios

Tenga en cuenta que este nodo puede tener elementos secundarios si se quita un subárbol del documento. También puede ocurrir que el nodo aparezca como "desconectado" y se le permita realizar consultas en sentido descendente desde este. Sin embargo, al intentar hacerlo en sentido ascendente, el nodo se mostrará como el único elemento presente.


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




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word en Office para iPad.|
|1.0|Agregado|
