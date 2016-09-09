
# Labs.Components.ActivityComponentInstance

 _**Hace referencia a:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Representa la instancia actual de un componente de actividad.

```
class ActivityComponentInstance extends Labs.ComponentInstance<Components.ActivityComponentAttempt>
```


## Propiedades


|**Nombre**|**Descripción**|
|:-----|:-----|
| `public var component: Components.IActivityComponentInstance`|El objeto subyacente [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md) que representa esta clase.|

## Métodos




### constructor

 `function constructor(component: Components.IActivityComponentInstance)`

Crea una nueva instancia de la clase [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md).

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _componente_|La instancia **IActivityComponentInstance** para crear esta clase a partir de esta.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ActivityComponentAttempt`

Crea una nueva instancia **ActivityComponentAttempt** e implementa el método abstracto que se define en la clase base

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _createAttemptResult_|El resultado de una acción de creación de intentos.|
