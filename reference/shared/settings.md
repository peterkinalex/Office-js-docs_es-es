
# <a name="settings-object"></a>Objeto Settings
Representa la configuración personalizada de un complemento de panel de tareas o de contenido que se almacena en el documento host como pares de nombre y valor.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint y Word|
|**Disponible en el [conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Configuración|
|**Modificado por última vez en**|1.1|

```
Office.context.document.settings
```


## <a name="members"></a>Miembros


**Métodos**

|||
|:-----|:-----|
|Nombre|Descripción|
|[addHandlerAsync](../../reference/shared/settings.addhandlerasync.md)|Agrega un controlador de eventos para el evento **settingsChanged**.|
|[get](../../reference/shared/settings.get.md)|Recupera la configuración especificada.|
|[refreshAsync](../../reference/shared/settings.refreshasync.md)|Lee toda la configuración que se conserva en el documento y actualiza la copia de esa configuración del complemento que se conserva en la memoria.|
|[remove](../../reference/shared/settings.remove.md)|Elimina la configuración especificada.|
|[removeHandlerAsync](../../reference/shared/settings.removehandlerasync.md)|Elimina un controlador de eventos para el evento **settingsChanged**.|
|[saveAsync](../../reference/shared/settings.saveasync.md)|Guarda la configuración.|
|[set](../../reference/shared/settings.set.md)|Define o crea la configuración especificada.|

**Eventos**


|**Nombre**|**Descripción**|
|:-----|:-----|
|[settingsChanged](../../reference/shared/settings.settingschangedevent.md)|Se produce cuando se cambia una configuración.|

## <a name="remarks"></a>Comentarios

La configuración que se crea con los métodos del objeto **Settings** se guarda por complemento y por documento. Es decir, solo está disponible para el complemento que la creó y solo desde el documento en el que se guarda.

El nombre de una configuración es una **string**, mientras que el valor puede ser **string**, **number**, **boolean**, **null**, **object** o **array**.

El objeto **Settings** se carga automáticamente como parte del objeto [Document](../../reference/shared/document.md) y está disponible al llamar a la propiedad [settings](../../reference/shared/document.settings.md) de dicho elemento cuando se activa el complemento. El desarrollador es responsable de llamar al método [saveAsync](../../reference/shared/settings.saveasync.md) después de agregar o suprimir la configuración para guardar la configuración en el documento.


## <a name="support-details"></a>Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos**|Configuración|
|**Tipos de complementos**|Contenido, panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## <a name="support-history"></a>Historial de compatibilidad

|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|Para los métodos **addHandlerAsync** y **removeHandlerAsync**, se agregó compatibilidad para agregar y quitar controladores de eventos para el evento en los complementos de contenido para Access. Para los métodos **get**, **refreshAsync**, **remove**, **saveAsync** y **set**, se agregó compatibilidad para la configuración personalizada en los complementos de contenido para Access.|
|1.0|Agregado|
