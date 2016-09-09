

# Objeto Office
Representa una instancia de un complemento, que proporciona acceso a objetos de nivel superior de la API.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project y Word|
|**Modificado por última vez en**|1.1|

```js
Office
```


## Miembros


**Propiedades**

|||
|:-----|:-----|
|Nombre|Descripción|
|[context](../../reference/shared/office.context.md)|Obtiene el objeto Context que representa el entorno en tiempo de ejecución del complemento y proporciona acceso a los objetos de nivel superior de la API.|
|[cast.item](../../reference/shared/office.cast.item.md)|Proporciona IntelliSense en Visual Studio específico para mensajes y citas en modo de redacción o lectura. <br/><br/><blockquote>**Nota** Aplicable únicamente en tiempo de diseño al desarrollar complementos de Outlook en Visual Studio. </blockquote>|

**Métodos**

|||
|:-----|:-----|
|Nombre|Descripción|
|[select](../../reference/shared/office.select.md)|Crea una promesa de devolver un enlace basado en la cadena de selector transferida.|
|[useShortNamespace](../../reference/shared/office.useshortnamespace.md)|Activa y desactiva el alias de **Office** para el espacio de nombres **Microsoft.Office.WebExtension** completo.|

**Eventos**

|||
|:-----|:-----|
|Nombre|Descripción|
|[initialize](../../reference/shared/office.initialize.md)|Ocurre cuando se carga el entorno en tiempo de ejecución y el complemento está preparado para empezar a interactuar con la aplicación y el documento alojado.|

## Comentarios

El objeto **Office** permite al desarrollador implementar una función de devolución de llamada para el evento Initialize y proporciona acceso al objeto [Context](../../reference/shared/context.md).


## Detalles de compatibilidad


Una Y mayúscula en la siguiente matriz indica que este objeto es compatible con la aplicación host de Office correspondiente. Una celda vacía indica que la aplicación host no admite este objeto.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office para escritorio de Windows**|**Office Online (en el explorador)**|**Office para iPad**|**OWA para dispositivos**|**Outlook para Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||v||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Project**|v|||||
|**Word**|v|v|v|||

|||
|:-----|:-----|
|**Tipos de complementos**|Contenido, Outlook y panel de tareas|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad


|**Versión**|**Cambios**|
|:-----|:-----|
|1.1|Se ha agregado compatibilidad para Excel, PowerPoint y Word en Office para iPad.|
|1.1|<ul><li>Para <a href="6c4b2c16-d4fb-4ecf-b72c-1e33b205daaf.htm">context</a>, se ha agregado compatibilidad para obtener el contexto en tiempo de ejecución con complementos de contenido para Acess.</p></li><li><p>Para <a href="23aeb136-da1f-4127-a798-99dc27bc4dae.htm">select</a>, se ha agregado compatibilidad para seleccionar enlaces de tabla con complementos de contenido para Acess.</li><li>Para <a href="9a4d5c7d-fcc4-4e8f-bef2-f2a8d8b4ae00.htm">useShortNamespace</a>, se ha agregado compatibilidad con complementos de contenido para Access.</li><li>Para <a href="727adf79-a0b5-48d2-99c7-6642c2c334fc.htm">initialize</a>, se ha agregado compatibilidad para inicialización en complementos de contenido para Access.</li></ul>|
|1.0|Agregado|

