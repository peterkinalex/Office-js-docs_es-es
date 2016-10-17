

# <a name="settings.settingschanged-event"></a>Evento Settings.settingsChanged
Se produce cuando la copia en memoria del contenedor de propiedades de la configuración se guarda en el documento con el método [Settings.saveAsync](../../reference/shared/settings.saveasync.md).

|||
|:-----|:-----|
|**Hosts:**|Excel |
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Configuración|
|**Modificado por última vez en**|1.0|

```js
Office.EventType.SettingsChanged
```


## <a name="remarks"></a>Comentarios

Para agregar un controlador de eventos para el evento **settingsChanged**, use el método [addHandlerAsync](../../reference/shared/settings.addhandlerasync.md) del objeto **Settings**.

El evento **settingsChanged** se desencadena solo cuando el script del complemento llama al método **Settings.saveAsync** para conservar en el archivo de documento la copia de la configuración que se encuentra en la memoria. El evento **settingsChanged** no se desencadena cuando se llama a los métodos [Settings.set](../../reference/shared/settings.set.md) o [Settings.remove](../../reference/shared/settings.remove.md).

El evento **settingsChanged** se ha diseñado para controlar los posibles conflictos que pueden originarse cuando varios usuarios intentan guardar la configuración al mismo tiempo mientras el complemento se usa en un documento compartido (con coautoría).


 >**Importante**: el código del complemento puede registrar un controlador para el evento **settingsChanged** cuando el complemento se ejecuta con cualquier cliente de Excel, pero el evento se activará solo si el complemento se carga con una hoja de cálculo que esté abierta en Excel Online _y_ más de un usuario esté editándola (coautoría). Por lo tanto, el evento **settingsChanged** solo se admite en Excel Online en escenarios de coautoría.


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este evento es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este evento.

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
