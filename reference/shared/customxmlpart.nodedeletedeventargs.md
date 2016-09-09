
# Objeto NodeDeletedEventArgs
Proporciona información sobre el nodo eliminado que generó el evento [dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md).

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Agregado en**|1.1|

```
NodeDeletedEventArgs
```


## Miembros


**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[isUndoRedo](../../reference/shared/customxmlpart.isundoredo.md)|Obtiene si el nodo se ha suprimido como parte de una acción Deshacer/Rehacer del usuario.|
|[oldNextSibling](../../reference/shared/customxmlpart.oldnextsibling.md)|Obtiene el siguiente elemento del mismo nivel antiguo del nodo que se acaba de suprimir del objeto **CustomXMLPart**.|
|[oldNode](../../reference/shared/customxmlpart.oldnode.md)|Obtiene el nodo que se acaba de eliminar del objeto **CustomXmlPart**.|

## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

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




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Word en Office para iPad.|
|1.0|Agregado|
