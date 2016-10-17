
# <a name="binding-object"></a>Objeto Binding
Una clase abstracta que representa un enlace a una sección del documento.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel y Word|
|**Disponible en los [conjuntos de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBinding, TableBinding, TextBinding|
|**Modificado por última vez en TableBinding**|1.1|

```js
Office.context.document.bindings.getByIdAsync(id);
```

## <a name="members"></a>Miembros


**Objetos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[MatrixBinding](../../reference/shared/binding.matrixbinding.md)|Representa un enlace en dos dimensiones de filas y columnas.|
|[TableBinding](../../reference/shared/binding.tablebinding.md)|Representa un enlace en dos dimensiones de filas y columnas, que puede llevar o no encabezados.|
|[TextBinding](../../reference/shared/binding.textbinding.md)|Representa una selección de texto enlazado en el documento.|

**Propiedades**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[document](../../reference/shared/binding.document.md)|Obtiene el objeto **Document** que se asocia con el enlace.|
|[id](../../reference/shared/binding.id.md)|Obtiene el identificador del objeto.|
|[type](../../reference/shared/binding.type.md)|Obtiene el tipo del enlace.|

**Métodos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/binding.addhandlerasync.md)|Agrega un controlador al enlace para el tipo de evento especificado.|
|[getDataAsync](../../reference/shared/binding.getdataasync.md)|Devuelve los datos que contiene el enlace.|
|[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|Quita del enlace el controlador que se especifica para el tipo de evento determinado.|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|Escribe datos en la sección enlazada del documento que representa el objeto de enlace que se ha especificado.|
|[TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|Establece o actualiza el formato de los elementos y datos especificados en la tabla enlazada.|

**Eventos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[bindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md)|Se produce al cambiar los datos en el enlace.|
|[bindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md)|Se produce al cambiar la selección en el enlace.|

## <a name="remarks"></a>Comentarios

El objeto **Binding** expone la funcionalidad que poseen todos los enlaces, independientemente de su tipo.

El objeto **Binding** nunca se llama de forma directa. Es la clase primaria abstracta de los objetos que representa cada tipo de enlace: [MatrixBinding](../../reference/shared/binding.matrixbinding.md), [TableBinding](../../reference/shared/binding.tablebinding.md) o [TextBinding](../../reference/shared/binding.textbinding.md). Estos tres objetos heredan los métodos **getDataAsync** y **setDataAsync** del objeto **Binding**, que habilitan al usuario a interactuar con los datos del enlace. También heredan las propiedades **id** y **type** para realizar consultas de estos valores de propiedad. Asimismo, los objetos **MatrixBinding** y **TableBinding** exponen métodos adicionales para las características específicas de matrices y tablas, como contar el número de filas y columnas.


## <a name="support-details"></a>Detalles de compatibilidad


La compatibilidad para cada miembro de API del objeto **Binding** difiere entre aplicaciones host de Office. Consulte la sección "Detalles de compatibilidad" del tema de cada miembro para obtener información de compatibilidad de host.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|MatrixBinding, TableBinding, TextBinding|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|
