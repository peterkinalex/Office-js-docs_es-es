
# Propiedad Context.roamingSettings
Obtiene un objeto que representa la configuración personalizada o el estado de un complemento de Outlook que se guardó en el buzón de correo de un usuario.

|||
|:-----|:-----|
|**Hosts:**|Outlook|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Buzón|
|**Modificado por última vez en**|1,0|

```
var appSettings = office.context.roamingSettings;
```


## Valor devuelto

Un objeto [RoamingSettings](http://msdn.microsoft.com/library/cf21bb08-7274-4ad6-ae9e-b2c12f92abc9%28Office.15%29.aspx).


## Comentarios

El objeto **RoamingSettings** le permite almacenar y tener acceso a datos para un complemento de Outlook almacenado en el buzón de correo de un usuario para que el complemento pueda obtener acceso a ellos cuando se ejecute desde cualquier aplicación de cliente host.


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este método es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este método.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Outlook para Mac**|
|:-----|:-----|:-----|:-----|
|**Outlook**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|Buzón|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Outlook|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad



****


|**Versión**|**Cambios**|
|:-----|:-----|
|1,0|Agregado|
