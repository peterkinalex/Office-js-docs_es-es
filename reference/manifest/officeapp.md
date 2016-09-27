
# Elemento OfficeApp
El elemento raíz del manifiesto de un complemento de Office.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## Sintaxis:


```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```


## Forma parte de:

 _ninguno_


## Debe contener:



|**Elemento**|**Contenido**|**Correo**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Id](../../reference/manifest/id.md)|x|x|x|
|[Versión](../../reference/manifest/version.md)|x|x|x|
|[ProviderName](../../reference/manifest/providername.md)|x|x|x|
|[DefaultLocale](../../reference/manifest/defaultlocale.md)|x|x|x|
|[DefaultSettings](../../reference/manifest/defaultsettings.md)|x|x|x|
|[DisplayName](../../reference/manifest/displayname.md)|x|x|x|
|[Descripción](../../reference/manifest/description.md)|x|x|x|
|[FormSettings](../../reference/manifest/formsettings.md)||x||
|[Permisos](../../reference/manifest/permissions.md)|x||x|
|[Regla](../../reference/manifest/rule.md)||x||

## Puede contener:



|**Elemento**|**Contenido**|**Correo**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](../../reference/manifest/alternateid.md)|x|x|x|
|[IconUrl](../../reference/manifest/iconurl.md)|x|x|x|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|x|x|x|
|[SupportUrl](../../reference/manifest/supporturl.md)|x|x|x|
|[AppDomains](../../reference/manifest/appdomains.md)|x|x|x|
|[Hosts](../../reference/manifest/hosts.md)|x|x|x|
|[Requisitos](../../reference/manifest/requirements.md)|x|x|x|
|[AllowSnapshot](../../reference/manifest/allowsnapshot.md)|x|||
|[Permisos](../../reference/manifest/permissions.md)||x||
|[DisableEntityHighlighting](../../reference/manifest/disableentityhighlighting.md)||x||
|[Dictionary](../../reference/manifest/dictionary.md)|||x|
|[VersionOverrides](../../reference/manifest/versionoverrides.md)|X|X|X|

## Atributos


|||
|:-----|:-----|
|xmlns|Define el esquema de versión y el espacio de nombres del manifiesto del complemento de Office. Este atributo debe establecerse siempre en `"http://schemas.microsoft.com/office/appforoffice/1.1"`|
|xmlns:xsi|Define la instancia de esquema XML. Este atributo debe establecerse siempre en `"http://www.w3.org/2001/XMLSchema-instance"`|
|xsi:type|Define el tipo de complemento de Office. Este atributo debe establecerse en uno de los siguientes valores: `"ContentApp"`, `"MailApp"` o `"TaskPaneApp"`|
