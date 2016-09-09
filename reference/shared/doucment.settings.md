
# Objeto Settings
Representa la configuración personalizada de un complemento de panel de tareas o de contenido que se almacena en el documento host como pares de nombre y valor.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint y Word|
|**Disponible en [el conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Configuración|
|**Modificado por última vez en**|1.1|

```
Office.context.document.settings
```


## Miembros


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


|**Name**|**Descripción**|
|:-----|:-----|
|[settingsChanged](../../reference/shared/settings.settingschangedevent.md)|Ocurre cuando se cambia una configuración.|

## Comentarios

La configuración que se crea con los métodos del objeto **Settings** se guarda por complemento y por documento. Es decir, solo está disponible para el complemento que la creó y solo desde el documento en el que se guarda.

El nombre de una configuración es una **string**, mientras que el valor puede ser **string**, **number**, **boolean**, **null**, **object** o **array**.

El objeto **Settings** se carga automáticamente como parte del objeto [Document](../../reference/shared/document.md) y está disponible al llamar a la propiedad [settings](../../reference/shared/document.settings.md) de dicho elemento cuando se activa el complemento. El desarrollador es responsable de llamar al método [saveAsync](../../reference/shared/settings.saveasync.md) después de agregar o suprimir la configuración para guardar la configuración en el documento.


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|v|
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v|v|

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|Configuración|
|**Tipos de complementos**|Panel de tareas y contenido|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Para los métodos <a href="7c4780cf-a779-4ac9-a362-c0bacae64a96.htm">addHandlerAsync</a> y <a href="735a255b-2a86-4b43-b1fa-e2a305815615.htm">removeHandlerAsync</a>, se ha agregado compatibilidad para agregar y quitar controladores de eventos para el evento <span class="keyword">SettingsChanged</span> en los complementos de contenido para Access. </p></li><li><p>Para los métodos <a href="aeac06dd-994e-4235-b208-1bd117395296.htm">get</a>, <a href="53a52c47-24b4-4d2d-b840-fe1b242cd795.htm">refreshAsync</a>, <a href="a92446bf-de65-45bd-8412-36ea8e77c5a2.htm">remove</a>, <a href="7147c221-937c-477c-98a6-f59d6200c27b.htm">saveAsync</a> y <a href="4e2c9758-953e-41e8-aca6-d8daf764a584.htm">set</a>, se ha agregado compatibilidad para la configuración personalizada en los complementos de contenido para Access.</p></li></ul>|
|1.0|Agregado|

