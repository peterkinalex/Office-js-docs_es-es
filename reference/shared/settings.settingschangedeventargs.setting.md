

# <a name="settingschangedeventargs.settings-property"></a>Propiedad SettingsChangedEventArgs.settings
Obtiene un objeto **Settings** que representa la configuración que generó el evento **settingsChanged**.

|||
|:-----|:-----|
|**Hosts:**|Excel|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Configuración|
|**Modificado por última vez en**|1.0|

```js
var mySettings = eventArgsObj.settings;
```


## <a name="return-value"></a>Valor devuelto

Objeto [Settings](../../reference/shared/document.settings.md) que representa la configuración que generó el evento [settingsChanged](../../reference/shared/settings.settingschangedevent.md).


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que esta propiedad es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||v||

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|Configuración|
|**Nivel de permisos mínimo**|[Restringido](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.0|Agregado|
