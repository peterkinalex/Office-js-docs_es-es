

# event.completed
La devolución de llamada que invoca el complemento para permitir que Outlook sepa que se ha realizado la operación.

****

|||
|:-----|:-----|
|**Hosts:** Outlook|**Tipo de complemento:** Outlook|
|**Disponible en los [conjuntos de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Buzón|
|**Última modificación en Buzón**|1.3|
|**Modos de Outlook aplicables**|Lectura y redacción|



```js
event.completed();
```


## Parámetros

Ninguno


## Detalles de compatibilidad


Una Y mayúscula en la siguiente tabla indica que esta propiedad es compatible con la aplicación host de Outlook correspondiente. Una celda vacía indica que la aplicación host de Outlook no es compatible con esta propiedad.

Para obtener más información sobre los requisitos de servidor y aplicación host de Office, consulte [Requisitos para ejecutar complementos de Office](../../docs/overview/requirements-for-running-office-add-ins.md).

 **Importante:** los comandos de complementos y las API asociadas de momento solo funcionan en Outlook en [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) para el escritorio de Windows.


**Hosts compatibles, por plataforma**


| |**Office para escritorio de Windows**|**Office Online (en el explorador)**|**OWA para dispositivos**|
|:-----|:-----|:-----|:-----|
|**Outlook**|v|||

|||
|:-----|:-----|
|**Disponible en los conjuntos de requisitos **|Buzón|
|**Nivel de permisos mínimo**|[ReadWriteItem](../../docs/outlook/understanding-outlook-add-in-permissions.md)|
|**Tipos de complementos**|Outlook|
|**Biblioteca**|Office.js|
|**Espacio de nombres**|Office|

## Historial de compatibilidad




|**Versión**|**Cambios**|
|:-----|:-----|
|1.3|Agregado|
