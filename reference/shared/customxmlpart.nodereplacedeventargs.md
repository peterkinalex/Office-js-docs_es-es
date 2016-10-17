
# <a name="nodereplacedeventargs-object"></a>Objeto NodeReplacedEventArgs
Proporciona información sobre el nodo reemplazado que generó el evento [dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md).

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Modificado por última vez en**|1.1|

```
NodeReplacedEventArgs
```


## <a name="members"></a>Miembros


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[isUndoRedo](../../reference/shared/customxmlpart.isundoredo.md)|Determina si el nodo reemplazado se ha insertado como parte de una operación de deshacer o rehacer del usuario.|
|[newNode](../../reference/shared/customxmlpart.newnode.md)|Determina el nuevo nodo.|
|[oldNode](../../reference/shared/customxmlpart.oldnode.md)|Determina el nodo antiguo (reemplazado).|

## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

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
