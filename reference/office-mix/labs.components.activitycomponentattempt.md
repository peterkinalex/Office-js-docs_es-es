
# <a name="labs.components.activitycomponentattempt"></a>Labs.Components.ActivityComponentAttempt

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Representa un intento de completar un componente de actividad.

```
class Permissions
```


## <a name="methods"></a>Métodos




### <a name="constructor"></a>constructor

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Crea una nueva instancia de la clase **ActivityComponentAttempt**.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _labs_|Instancias de laboratorio ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) asociadas al componente.|
| _componentId_|Identificador del componente asociado al intento.|
| _attemptId_|Identificador del intento.|
| _values_|Valores, si los hubiera, asociados al componente.|

### <a name="complete"></a>complete

 `public function complete(callback: Labs.Core.ILabCallback<void>): void`

Indicador de que la actividad se ha completado.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _callback_|Función de devolución de llamada que se invoca una vez que se ha completado la actividad.|

### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

Función que se ejecuta en las acciones que se recuperan para un intento determinado y después rellena el estado del laboratorio.

 **Parámetros**


|**Nombre**|**Descripción**|
|:-----|:-----|
| _action_|La instancia de acción ([Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md)).|
