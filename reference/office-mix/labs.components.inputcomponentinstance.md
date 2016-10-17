
# <a name="labs.components.inputcomponentinstance"></a>Labs.Components.InputComponentInstance

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Representa una instancia de un componente de entrada.

```
class InputComponentInstance extends Labs.ComponentInstance<Components.InputComponentAttempt>
```


## <a name="properties"></a>Propiedades


|Propiedad|Descripción|
|:-----|:-----|
| `public var component: Components.IInputComponentInstance`|El objeto subyacente [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) que representa esta clase.|

## <a name="methods"></a>Métodos




### <a name="constructor"></a>constructor

 `function constructor(component: Components.IInputComponentInstance)`

Crea una nueva instancia [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md).

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _component_|El objeto [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) desde el que se crea esta clase.|

### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.InputComponentAttempt`

Crea un nuevo objeto [Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md). Implementa el método abstracto que se define en la clase base.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _createAttemptResult_|El resultado de una acción de creación de intentos.|
