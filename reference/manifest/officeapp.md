
# <a name="officeapp-element"></a>Elemento OfficeApp
El elemento raíz del manifiesto de un complemento de Office.

 **Tipo de complemento:** Contenido, panel de tareas, correo


## <a name="syntax:"></a>Sintaxis:


```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```


## <a name="contained-in:"></a>Forma parte de:

 _ninguno_


## <a name="must-contain:"></a>Debe contener:



|**Elemento**|**Contenido**|**Correo**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Id](../../reference/manifest/id.md)|x|x|x|
|[Version](../../reference/manifest/version.md)|x|x|x|
|[ProviderName](../../reference/manifest/providername.md)|x|x|x|
|[DefaultLocale](../../reference/manifest/defaultlocale.md)|x|x|x|
|[DefaultSettings](../../reference/manifest/defaultsettings.md)|x|x|x|
|[DisplayName](../../reference/manifest/displayname.md)|x|x|x|
|[Description](../../reference/manifest/description.md)|x|x|x|
|[FormSettings](../../reference/manifest/formsettings.md)||x||
|[Permissions](../../reference/manifest/permissions.md)|x||x|
|[Rule](../../reference/manifest/rule.md)||x||

## <a name="can-contain:"></a>Puede contener:



|**Elemento**|**Contenido**|**Correo**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](../../reference/manifest/alternateid.md)|x|x|x|
|[IconUrl](../../reference/manifest/iconurl.md)|x|x|x|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|x|x|x|
|[SupportUrl](../../reference/manifest/supporturl.md)|x|x|x|
|[AppDomains](../../reference/manifest/appdomains.md)|x|x|x|
|[Hosts](../../reference/manifest/hosts.md)|x|x|x|
|[Requirements](../../reference/manifest/requirements.md)|x|x|x|
|[AllowSnapshot](../../reference/manifest/allowsnapshot.md)|x|||
|[Permissions](../../reference/manifest/permissions.md)||x||
|[DisableEntityHighlighting](../../reference/manifest/disableentityhighlighting.md)||x||
|[Dictionary](../../reference/manifest/dictionary.md)|||x|
|[VersionOverrides](../../reference/manifest/versionoverrides.md)|X|X|X|

## <a name="attributes"></a>Atributos


|||
|:-----|:-----|
|xmlns|Define el esquema de versión y el espacio de nombres del manifiesto del complemento de Office. Este atributo debe establecerse siempre en `"http://schemas.microsoft.com/office/appforoffice/1.1"`|
|xmlns:xsi|Define la instancia de esquema XML. Este atributo debe establecerse siempre en `"http://www.w3.org/2001/XMLSchema-instance"`|
|xsi:type|Define el tipo de complemento de Office. Este atributo debe establecerse en uno de los siguientes valores: `"ContentApp"`, `"MailApp"` o `"TaskPaneApp"`|
