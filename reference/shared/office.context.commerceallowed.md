
# <a name="context.commerceallowed-property"></a>Propiedad Context.commerceAllowed
Obtiene información sobre si el complemento se está ejecutando en una plataforma que admite vínculos a sistemas de pago externos.

|||
|:-----|:-----|
|**Hosts:**|Excel y Word|
|**Modificado por última vez en**|1.1|

```
var allowCommerce = Office.context.commerceAllowed;
```


## <a name="return-value"></a>Valor devuelto

Devuelve **True** si los desarrolladores pueden mostrar la IU de venta o actualización del complemento en esa plataforma; de lo contrario, devuelve **False**.


## <a name="remarks"></a>Comentarios

El App Store de iOS no admite las aplicaciones con complementos que incluyan vínculos a sistemas de pago adicionales. Sin embargo, los complementos de Office que se ejecutan en el escritorio de Windows o en Office Online en el explorador sí permiten esos vínculos. Si desea que la IU del complemento incluya un vínculo a un sistema de pago externo en plataformas que no sean iOS, puede usar la propiedad **commerceAllowed** para controlar cuándo se muestra ese vínculo.


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para iPad**|
|:-----|:-----|
|**Excel**|v|
|**PowerPoint**||
|**Word**|v|

|||
|:-----|:-----|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Agregado.|
