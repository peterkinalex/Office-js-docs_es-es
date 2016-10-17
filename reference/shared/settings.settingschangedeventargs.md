# <a name="settings.settingschangedeventargs-object"></a>Objeto Settings.settingschangedeventargs
Proporciona información sobre la configuración que generó el evento [settingsChanged](settings.settingschangedevent.md).

|||
|:-----|:-----|
|**Hosts:**|Access y Excel |
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Configuración|
|**Modificado por última vez en**|1.0|

```js
Office.EventType.SettingsChanged
```

## <a name="members"></a>Miembros

**Propiedades**

|**Nombre**|**Descripción**|
|:-----|:-----|
|**[settings](settings.settingschangedeventargs.setting.md)**|Obtiene un objeto **Settings** que representa la configuración que generó el evento settingsChanged.|
|**[type](settings.settingschangedeventargs.type.md)**|Obtiene un valor de la enumeración **EventType** que identifica el tipo de evento que se generó.|

## <a name="remarks"></a>Comentarios

Para agregar un controlador de eventos para el evento **settingsChanged**, use el método [addHandlerAsync](settings.addhandlerasync.md) del objeto **Settings**.

El evento **settingsChanged** se desencadena solo cuando el script del complemento llama al método **Settings.saveAsync** para conservar en el archivo de documento la copia de la configuración que se encuentra en la memoria. El evento **settingsChanged** no se desencadena cuando se llama a los métodos [Settings.set](settings.set.md) o [Settings.remove](settings.remove.md).

El evento **settingsChanged** se ha diseñado para controlar los posibles conflictos que pueden originarse cuando varios usuarios intentan guardar la configuración al mismo tiempo mientras el complemento se usa en un documento compartido (con coautoría).


 >**Importante**: el código del complemento puede registrar un controlador para el evento **settingsChanged** cuando el complemento se ejecuta con cualquier cliente de Excel, pero el evento se activará solo si el complemento se carga con una hoja de cálculo que esté abierta en Excel Online _y_ más de un usuario esté editándola (coautoría). Por lo tanto, el evento **settingsChanged** solo se admite en Excel Online en escenarios de coautoría.



## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||v||


|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|Configuración|
|**Nivel de permisos mínimo**|Restringido|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad

|**Versión**|**Cambios**|
|:-----|:-----|
|1.0|Agregado|
