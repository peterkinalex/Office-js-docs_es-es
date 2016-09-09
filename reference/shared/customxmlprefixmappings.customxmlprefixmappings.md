
# Objeto CustomXmlPrefixMappings
Representa una colección de asignaciones personalizadas de prefijo de espacio de nombres.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Modificado por última vez en**|1.1|

```
CustomXmlPrefixMappings
```


## Miembros


**Métodos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[addNamespaceAsync](../../reference/shared/customxmlprefixmappings.addnamespaceasync.md)|Agrega de forma asincrónica un prefijo a una asignación de espacio de nombres para usarla cuando se consulte un elemento.|
|[getNamespaceAsync](../../reference/shared/customxmlprefixmappings.getnamespaceasync.md)|Obtiene de forma asíncrona el espacio de nombres asignado al prefijo especificado.|
|[getPrefixAsync](../../reference/shared/customxmlprefixmappings.getprefixasync.md)|Obtiene de forma asincrónica el prefijo para el espacio de nombres que se ha especificado.|

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
