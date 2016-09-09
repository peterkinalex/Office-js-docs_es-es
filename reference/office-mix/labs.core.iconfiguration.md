
# Labs.Core.IConfiguration

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Estructura de datos de una configuración de laboratorio.

```
interface IConfiguration extends Core.IUserData
```


## Propiedades


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|Versión de la aplicación asociada a esta configuración.|
| `components: Core.IComponent[]`|Componentes incluidos en el laboratorio.|
| `name: string`|El nombre del laboratorio.|
| `timeline: Core.ITimelineConfiguration`|La configuración de la escala de tiempo del laboratorio.|
| `analytics: Core.IAnalyticsConfiguration`|La configuración de análisis del laboratorio.|
