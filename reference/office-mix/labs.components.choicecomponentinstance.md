
# <a name="labs.components.choicecomponentinstance"></a>Labs.Components.ChoiceComponentInstance

 _**Hace referencia a:** apps para Office | Complementos de Office | Office Mix | PowerPoint_

Representa una instancia de un componente de elección.

```
class ChoiceComponentInstance extends Labs.ComponentInstance<Components.ChoiceComponentAttempt>
```


## <a name="properties"></a>Propiedades


|Propiedad|Descripción|
|:-----|:-----|
| `public var component: Components.IChoiceComponentInstance`|El objeto subyacente [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) que representa esta clase.|

## <a name="methods"></a>Métodos




### <a name="constructor"></a>constructor

 `function constructor(component: Components.IChoiceComponentInstance)`

Crea una nueva instancia de la clase **ChoiceComponentInstance**.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _component_|El objeto [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) desde el que se crea esta clase.|

### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ChoiceComponentAttempt`

Crea una nueva instancia **ChoiceComponentAttempt** e implementa el método abstracto que se define en la clase base.

 **Parámetros**


|Parámetro|Descripción|
|:-----|:-----|
| _createAttemptResult_|El resultado de la acción de creación de intentos.|
