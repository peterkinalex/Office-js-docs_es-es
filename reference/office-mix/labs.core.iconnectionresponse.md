
# Labs.Core.IConnectionResponse

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Información de respuesta que se devuelve desde una llamada de conexión.

```
interface IConnectionResponse
```


## Properties


|||
|:-----|:-----|
| `initializationInfo: Core.IConfigurationInfo`|Inicialización de información de configuración o **null** si la aplicación no se ha inicializado.|
| `mode: Core.LabMode`|El modo en el que el laboratorio está ejecutándose actualmente.|
| `hostVersion: Core.IVersion`|Información de la versión ([Labs.Core.IVersion](../../reference/office-mix/labs.core.iversion.md)) para el servidor.|
| `userInfo: Core.IUserInfo`|Información sobre el usuario ([Labs.Core.IUserInfo](../../reference/office-mix/labs.core.iuserinfo.md)).|
