
# <a name="labs.core.iconfigurationinstance"></a>Labs.Core.IConfigurationInstance

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Clase base para las instancias de una configuración de laboratorio. Una instancia es una creación de instancia de una configuración para un usuario determinado y contiene una vista traducida de la configuración de una ejecución particular del laboratorio. Esta vista puede excluir información oculta (por ejemplo, sugerencias y respuestas) y también contiene identificadores para identificar las diversas instancias.

```
interface IConfigurationInstance extends Core.IUserData
```


## <a name="properties"></a>Propiedades


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|Versión del laboratorio asociado a esta configuración.|
| `components: Core.IComponentInstance[]`|Componentes asociados al laboratorio.|
| `name: string`|Nombre del laboratorio.|
| `timeline: Core.ITimelineConfiguration`|Configuración de escala de tiempo para el laboratorio.|
