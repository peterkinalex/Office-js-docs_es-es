
# <a name="permissions-element"></a>Elemento Permissions
Especifica el nivel de acceso a la API para su complemento de Office; debe solicitar permisos según el principio de privilegios mínimos.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## <a name="syntax:"></a>Sintaxis:

Para complementos de paneles de tareas:


```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Para complementos de correo:




```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```


## <a name="contained-in:"></a>Forma parte de:

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## <a name="remarks"></a>Comentarios

Para obtener más información, consulte [Solicitar permisos para el uso de API en complementos de contenido y de panel de tareas](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) e [Información sobre los permisos de los complementos de Outlook](../../docs/outlook/understanding-outlook-add-in-permissions.md).

