
# Objeto CustomXmlParts
Representa una colección de objetos [CustomXMLPart](../../reference/shared/customxmlpart.customxmlpart.md).

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Modificado por última vez en**|1.1|

```
Office.context.document.customXmlParts
```


## Miembros


**Métodos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[addAsync](../../reference/shared/customxmlparts.addasync.md)|Agrega de forma asíncrona un nuevo elemento XML personalizado a un archivo.|
|[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|Obtiene de forma asíncrona un elemento XML personalizado por su id.|
|[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|Obtiene de forma asíncrona una matriz de elementos XML personalizados que coinciden con el espacio de nombres especificado.|

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
