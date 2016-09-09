
# Objeto TextBinding
Representa una selección de texto enlazado en el documento.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TextBindings|
|**Agregado en**|1,0|

```
TextBinding
```


## Comentarios

El objeto **TextBinding** hereda las propiedades [id](../../reference/shared/binding.id.md) y [type](../../reference/shared/binding.type.md), y los métodos [getDataAsync](../../reference/shared/binding.getdataasync.md) y [setDataAsync](../../reference/shared/binding.setdataasync.md) del objeto [Binding](../../reference/shared/binding.md). No implementa ninguna otra de sus propiedades ni métodos adicionales.


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hosts compatibles, por plataforma**


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|TextBindings|
|**Nivel de permisos mínimo**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel y Word en Office para iPad.|
|1.0|Agregado|
